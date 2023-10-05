export function logClear() {
  try {
    document.getElementById("message").innerHTML = "";
  } catch (e) {
    console.error(e);
  }
}

export function log(...args) {
  console.log(...args);
  let message = "";
  try {
    message = args.map((arg) => `<pre><code>${JSON.stringify(arg, null, "  ")}</code></pre>`).join("");
  } catch (e) {
    console.error(e);
  }
  try {
    document.getElementById("message").innerHTML = message + document.getElementById("message").innerHTML;
  } catch (e) {
    console.error(e);
  }
}

export function logKeys(label, obj) {
  try {
    console.log(keysRecursive(label, obj));
  } catch (e) {
    console.log([label, e.toString()]);
  }
}

// HELPERS: //

function shortValType(val) {
  let type = typeof val;
  switch (type) {
    case "object":
      if (Array.isArray(val)) return "[]";
      return "{}";
    case "function":
      return "()";
    case "string":
      return val.substring(0, 20) + (val.length > 20 ? "..." : "");
    case "number":
    case "boolean":
      return val;
    default:
      return type;
  }
}

function keysRecursive(key, val) {
  try {
    if (typeof val === "object") {
      let keys = [];
      for (let key in val) {
        if (!key) continue;
        if (key.substring(0, 1) === "_") continue;
        if (key.substring(0, 2) === "m_") continue;
        try {
          keys.push(key + " " + shortValType(val[key]));
        } catch (e) {
          keys.push(key);
        }
      }
      return [key, keys];
    } else {
      return [key, val.toJSON ? val.toJSON : val.toString()];
    }
  } catch (e) {
    return [key, val];
  }
}
