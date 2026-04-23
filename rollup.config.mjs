import typescript from "@rollup/plugin-typescript";
import { dts } from "rollup-plugin-dts";

const input = "src/index.ts";

export default [
  {
    input,
    external: ["jszip", "node:fs/promises"],
    output: [
      { file: "dist/index.js", format: "esm", sourcemap: true },
      { file: "dist/index.cjs", format: "cjs", sourcemap: true, exports: "named" }
    ],
    plugins: [typescript({ tsconfig: "./tsconfig.build.json" })]
  },
  {
    input,
    output: [{ file: "dist/index.d.ts", format: "es" }],
    plugins: [dts()]
  }
];
