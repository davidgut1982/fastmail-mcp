// Normalize incoming MCP tool args so cheap non-Anthropic models (Llama,
// Mistral, etc. via DeepInfra/LiteLLM) that emit snake_case can hit our
// camelCase schema without the SDK validator rejecting the call. We build
// a per-tool snake_case → camelCase map dynamically by walking each tool's
// declared inputSchema, then fix up incoming args before dispatch.
//
// Goals:
//   - Idempotent. Already-camelCase args pass through untouched.
//   - Conservative. Only keys that EXACTLY MATCH a known camelCase key after
//     snake-to-camel conversion are renamed. Unknown keys are left alone, so
//     a typo doesn't get silently "fixed" into something wrong.
//   - Recursive. Nested objects (e.g. participant entries, filter blocks) and
//     arrays of objects get the same treatment.
//   - Pure. The only side effect is a single console.error log line per rename
//     event, so the operator can grep stderr to see what the cheap model sent.
//   - Bounded log noise. The caller decides whether to log original-vs-
//     normalized args; this module just tracks whether anything changed.
//
// We deliberately skip schema relaxation (oneOf with snake_case+camelCase
// required keys) — that bloats the schema for the strong models. Server-side
// normalization keeps tool descriptions clean.

type JsonSchema = {
  type?: string | string[];
  properties?: Record<string, JsonSchema>;
  items?: JsonSchema | JsonSchema[];
  required?: string[];
  [k: string]: any;
};

// snake_case → camelCase (no-op if already camelCase).
// Only flips the segment immediately after each underscore. Leading/trailing
// underscores are preserved (defensive — shouldn't occur in our schemas).
function snakeToCamel(s: string): string {
  if (!s.includes('_')) return s;
  return s.replace(/_([a-zA-Z0-9])/g, (_match, ch: string) => ch.toUpperCase());
}

// True iff a key contains a snake_case underscore that could plausibly be
// converted (i.e. has at least one alphanumeric on each side).
function looksLikeSnakeCase(s: string): boolean {
  return /[a-zA-Z0-9]_[a-zA-Z0-9]/.test(s);
}

// Build a recursive map of "known camelCase keys" from a JSON Schema. The map
// shape mirrors the schema: leaves are `true` (terminal), nested objects hold
// child maps. Arrays-of-objects collapse into the items' map. We don't track
// types — we only care about which keys EXIST at each nesting level.
type KnownKeys = {
  // The set of camelCase keys legal at THIS object level.
  keys: Set<string>;
  // For each known key, the map for its value (if it's an object/array of
  // objects). Missing entry → leaf, no recursion needed.
  children: Map<string, KnownKeys>;
};

function buildKnownKeys(schema: JsonSchema | undefined): KnownKeys {
  const out: KnownKeys = { keys: new Set(), children: new Map() };
  if (!schema || typeof schema !== 'object') return out;

  // Object level: walk properties.
  if (schema.properties && typeof schema.properties === 'object') {
    for (const [propName, propSchema] of Object.entries(schema.properties)) {
      out.keys.add(propName);
      const childKnown = buildKnownKeysFor(propSchema as JsonSchema);
      if (childKnown.keys.size > 0) {
        out.children.set(propName, childKnown);
      }
    }
  }

  return out;
}

// Build the KnownKeys for a single property's schema. Handles object,
// array-of-object, and the rare oneOf/anyOf shapes by merging.
function buildKnownKeysFor(propSchema: JsonSchema): KnownKeys {
  if (!propSchema || typeof propSchema !== 'object') {
    return { keys: new Set(), children: new Map() };
  }

  // Direct object.
  if (propSchema.properties) {
    return buildKnownKeys(propSchema);
  }

  // Array — recurse into items.
  if (propSchema.type === 'array' && propSchema.items) {
    const items = Array.isArray(propSchema.items) ? propSchema.items[0] : propSchema.items;
    return buildKnownKeysFor(items as JsonSchema);
  }

  // oneOf / anyOf — merge the keys from all branches.
  const branches: JsonSchema[] = [];
  if (Array.isArray(propSchema.oneOf)) branches.push(...propSchema.oneOf);
  if (Array.isArray(propSchema.anyOf)) branches.push(...propSchema.anyOf);
  if (branches.length > 0) {
    const merged: KnownKeys = { keys: new Set(), children: new Map() };
    for (const b of branches) {
      const sub = buildKnownKeysFor(b);
      for (const k of sub.keys) merged.keys.add(k);
      for (const [k, v] of sub.children) merged.children.set(k, v);
    }
    return merged;
  }

  return { keys: new Set(), children: new Map() };
}

// Result of normalizing one args payload. `changed` is true iff at least one
// key was renamed; the dispatcher uses this to decide whether to log the
// before/after pair.
export type NormalizeResult = {
  args: Record<string, any>;
  changed: boolean;
  // List of "fromKey -> toKey" renames performed, with dotted paths if nested.
  renames: Array<{ from: string; to: string; path: string }>;
};

// Internal recursive worker. Logs each rename via console.error, so we don't
// have to thread a logger through.
function normalizeNode(
  node: any,
  known: KnownKeys,
  toolName: string,
  pathPrefix: string,
  renames: Array<{ from: string; to: string; path: string }>
): any {
  // Arrays: walk each element. Use `known` as-is — it represents the items'
  // schema (buildKnownKeysFor unwraps array container).
  if (Array.isArray(node)) {
    return node.map((el, i) =>
      normalizeNode(el, known, toolName, `${pathPrefix}[${i}]`, renames)
    );
  }

  // Plain objects: rename keys and recurse into values.
  if (node && typeof node === 'object' && node.constructor === Object) {
    const out: Record<string, any> = {};
    for (const [k, v] of Object.entries(node)) {
      let targetKey = k;

      // Only consider renaming if the key looks snake_case AND the camelCased
      // form is in the known set AND the snake_case form ITSELF isn't already
      // in the known set (which would mean the schema actually expects it).
      if (looksLikeSnakeCase(k) && !known.keys.has(k)) {
        const camel = snakeToCamel(k);
        if (camel !== k && known.keys.has(camel)) {
          targetKey = camel;
          const dottedPath = pathPrefix ? `${pathPrefix}.${k}` : k;
          renames.push({ from: k, to: camel, path: dottedPath });
          console.error(
            `[fastmail-mcp] param-normalize: ${k} → ${camel} in ${toolName}` +
              (pathPrefix ? ` (at ${pathPrefix})` : '')
          );
        }
      }

      // Recurse into the value with the child schema (if any).
      const childKnown = known.children.get(targetKey);
      if (childKnown) {
        const nextPath = pathPrefix ? `${pathPrefix}.${targetKey}` : targetKey;
        out[targetKey] = normalizeNode(v, childKnown, toolName, nextPath, renames);
      } else {
        out[targetKey] = v;
      }
    }
    return out;
  }

  // Primitive / null / undefined: unchanged.
  return node;
}

// Public entry point. Pass the args object and the tool's inputSchema; we
// return a (possibly new) args object plus a `changed` flag. Original args
// are never mutated.
export function normalizeToolArgs(
  args: Record<string, any> | undefined | null,
  inputSchema: JsonSchema | undefined,
  toolName: string
): NormalizeResult {
  if (!args || typeof args !== 'object') {
    return { args: args ?? {}, changed: false, renames: [] };
  }
  if (!inputSchema || typeof inputSchema !== 'object') {
    return { args, changed: false, renames: [] };
  }

  const known = buildKnownKeys(inputSchema);
  if (known.keys.size === 0) {
    // Schema declared no properties (e.g. {} input). Pass through.
    return { args, changed: false, renames: [] };
  }

  const renames: Array<{ from: string; to: string; path: string }> = [];
  const out = normalizeNode(args, known, toolName, '', renames);
  return {
    args: out as Record<string, any>,
    changed: renames.length > 0,
    renames,
  };
}

// Test-only helper: expose the pure key converter.
export const __test = { snakeToCamel, looksLikeSnakeCase, buildKnownKeys, buildKnownKeysFor };
