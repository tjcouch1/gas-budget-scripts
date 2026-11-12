// @ts-check
import eslint from "@eslint/js";
import prettier from "eslint-config-prettier";
import tseslint from "typescript-eslint";

export default tseslint.config(
  {
    ignores: [
      "**/dist/**",
      "**/coverage/**",
      "**/node_modules/**",
      "packages/budget-web/public/licenses.html",
    ],
  },
  eslint.configs.recommended,
  tseslint.configs.recommended,
  {
    rules: {
      "@typescript-eslint/no-unused-vars": [
        "error",
        { argsIgnorePattern: "^_", varsIgnorePattern: "^_" },
      ],
      eqeqeq: ["error", "always", { null: "ignore" }],
      // Comments contain verbatim sample emails with non-breaking spaces and BOMs
      "no-irregular-whitespace": ["error", { skipComments: true }],
      "prefer-const": "error",
    },
  },
  {
    // Build tooling and configs run in Node
    files: [
      "packages/*/scripts/**/*.ts",
      "*.config.{js,ts}",
      "packages/*/*.config.{js,ts}",
    ],
    languageOptions: { globals: { process: "readonly", console: "readonly" } },
  },
  prettier,
);
