// eslint.config.js
import { defineConfig } from "eslint/config";
import globals from "globals";
import tseslint from "typescript-eslint";
import sortDestructureKeys from "eslint-plugin-sort-destructure-keys";
import sortExports from "eslint-plugin-sort-exports";

export default defineConfig([
	...tseslint.configs.recommended,
	{
		files: ["**/*.{js,mjs,cjs,ts}"],
		ignores: ["build/**", "dist/**", "node_modules/**"],
		languageOptions: {
			globals: globals.node,
		},
		plugins: {
			"sort-destructure-keys": sortDestructureKeys,
			"sort-exports": sortExports,
		},
		rules: {
			semi: ["error", "always"],
			"comma-dangle": ["error", "always-multiline"],
			"sort-destructure-keys/sort-destructure-keys": ["error", { caseSensitive: true }],
			"sort-exports/sort-exports": ["error", { sortDir: "asc", sortExportKindFirst: "type" }],
			"no-multiple-empty-lines": ["error", { max: 1, maxEOF: 0 }],
		},
	},
]);
