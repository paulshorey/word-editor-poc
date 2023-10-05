/* global Word, require */

// Utility function that just loads a component.
// If you try to do this in the code, eslint is sad and will not eslint'e anything in that file.
// Using this utility helps with that problem.
export function contextLoad(context, controls) {
  context.load(controls);
}
