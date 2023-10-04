export default function keys(label, obj) {
  document.getElementById("message").innerHTML = `
    <h3>${label}</h3>
    <pre><code>${JSON.stringify(keysRecursive(obj), null, "  ")}</code></pre>
  `;
}

function keysRecursive(obj) {
  if (typeof obj === "object") {
    let keys = [];
    for (let key in obj) {
      if (!key) return;
      if (key.substring(0, 1) === "_") continue;
      if (key.substring(0, 2) === "m_") continue;
      if (typeof keys[key] === "object") {
        keys[key] = keysRecursive(obj[key]);
      } else {
        keys.push(key);
      }
    }
    return keys;
  } else {
    return obj.toString();
  }
}
