import { expect, test } from "@jest/globals";
import { add } from "../lib";

test("can call our function", async () => {
  const actual = add(1, 2);
  await expect(actual).toEqual(3);
});
