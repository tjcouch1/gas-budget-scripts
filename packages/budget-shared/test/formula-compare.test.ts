import { describe, expect, it } from "vitest";
import { compareFormulas } from "../src/index.ts";

describe("compareFormulas", () => {
  it.each([
    ["=1+3.23*2", "= 1 + 3.23 * 2"],
    ["=.5", "=0.5"],
    ["=1.50*2", "=1.5*2"],
    ["=1.0825*(4.23*2)", "=1.0825*(4.23*2)"],
    ["=45-R[1]C[0]-R[2]C[0]", "=45-R[1]C-R[2]C"],
    ["=45-c12-c13", "=45-C12-C13"],
    ["=45-$C$12", "=45-C12"],
    ["=(40+5)-C12", "=(40+5)-C12"],
  ])("%s and %s are the same formula", (a, b) => {
    expect(compareFormulas(a, b)).toBe("same");
    expect(compareFormulas(b, a)).toBe("same");
  });

  it.each([
    ["=1*1.09", "=1.09"],
    ["=1.0825*(5)", "=1.0825*(6)"],
    ["=1.0825*(5)", "=1.1*(5)"],
    ["=45-C12", "=45-C13"],
    ["=45-C12-C13", "=45-C12"],
    ["=(40+5)-C12", "=45-C12"],
    ["=45-R[1]C[0]", "=45-R[2]C[0]"],
    ["=46-R[1]C[0]", "=45-C12"],
    ["=45-R[1]C[0]-R[2]C[0]", "=45-C12"],
    ["=12", "=1+2"],
  ])("%s and %s differ", (a, b) => {
    expect(compareFormulas(a, b)).toBe("different");
    expect(compareFormulas(b, a)).toBe("different");
  });

  it("cannot tell A1 and R1C1 references apart without the cell's position", () => {
    expect(compareFormulas("=45-R[1]C[0]-R[2]C[0]", "=45-C12-C13")).toBe(
      "refNotationDiffers",
    );
    expect(compareFormulas("=(40+5)-C12", "=(40+5)-R[1]C")).toBe(
      "refNotationDiffers",
    );
  });
});
