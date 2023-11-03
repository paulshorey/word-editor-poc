export default function resetControl(
  control: {
    tag: string;
    color: string;
    appearance?: string;
    [key: string]: any;
  },
  appearance?: string
) {
  control.color = "#666666";
  // if using nested controls (to display status/type)
  if (control.tag === ":") {
    // inner control
    control.appearance = appearance || "Tags";
  } else {
    // outer control
    control.appearance = appearance || "BoundingBox";
  }
  control.cannotDelete = false;
  control.cannotEdit = false;
  control.styleBuiltIn = "Strong";
}
