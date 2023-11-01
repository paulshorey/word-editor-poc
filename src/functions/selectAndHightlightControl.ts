/* global console, setTimeout, Office, document, Word, require */

const debounceSelectedTag = {
  id: "",
};

export default async function selectAndHightlightControl(control: any, context: any): Promise<void> {
  control.load("id");
  await context.sync();
  // do not scroll to same control - will start an infinite loop!
  if (debounceSelectedTag.id === control.id) return;
  debounceSelectedTag.id = control.id;
  // if new control, then go ahead and scroll, select, highlight
  control.select("Select");
  control.load("color");
  await context.sync();
  // control.color = "#F5C027";
  control.color = "#08E5FF";
  setTimeout(async () => {
    await context.sync();
    control.color = "#666666";
    context.sync();
  }, 1000);
}
