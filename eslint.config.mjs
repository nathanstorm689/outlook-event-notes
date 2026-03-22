import tsparser from "@typescript-eslint/parser";
import obsidianmd from "eslint-plugin-obsidianmd";

export default [
	{
		files: ["main.ts"],
		plugins: { obsidianmd },
		languageOptions: {
			parser: tsparser,
			parserOptions: { project: "./tsconfig.json" },
		},
		rules: {
			...obsidianmd.configs.recommended,
			"obsidianmd/ui/sentence-case": ["error", {
				brands: ["Outlook", "Classic", "Microsoft"]
			}],
		},
	},
];
