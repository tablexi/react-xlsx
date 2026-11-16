/* tslint:disable */
/* eslint-disable */

export function normalizeRaster(source: string): string | undefined;

export function renderPicture(xml: string, source: string, width: number, height: number, colors: Function): string | undefined;

/**
 * Render a shape or connector's geometry, fill and outline. `styles` resolves
 * a theme style reference: it receives "fill", "line" or "effect", the
 * reference's index and its XML, and returns the theme style's XML with the
 * placeholder colour applied, or undefined.
 */
export function renderShape(properties_xml: string, style_xml: string, width: number, height: number, colors: Function, styles: Function): string | undefined;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly normalizeRaster: (a: number, b: number, c: number) => void;
    readonly renderPicture: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number) => void;
    readonly renderShape: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number) => void;
    readonly __wbindgen_export: (a: number, b: number) => number;
    readonly __wbindgen_export2: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_export3: (a: number) => void;
    readonly __wbindgen_add_to_stack_pointer: (a: number) => number;
    readonly __wbindgen_export4: (a: number, b: number, c: number) => void;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;
