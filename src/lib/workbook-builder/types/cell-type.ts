/**
 * Supported Excel cell types.
 *
 * s           - Shared string (reference to sharedStrings.xml). <v> contains index in sharedStrings
 * inlineStr   - Inline string. Wrap value in <is><t>value</t></is>
 * str         - Formula string result. Not for plain text, only formula result
 * b           - Boolean. <v> — 0 or 1
 * e           - Error. <v> — error code
 * n           - Number. No attribute t or t="n"
*/
export type CellType = "s" | "inlineStr" | "n" | "b" | "str" | "e";
