export default function log(...args) {
  document.getElementById("message").innerHTML = `
    ${args.map((arg) => `<pre><code>${JSON.stringify(arg, null, "  ")}</code></pre>`).join("")}
  `;
}
