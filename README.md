<img width="1594" alt="image" src="https://github.com/paulshorey/word-editor-poc/assets/7524065/0d2c468e-c5d8-4281-a4fc-de4b19ab1c66">

# Resources:

Manually convert Word file to base64 string, then paste into our add-in app text area:
https://products.aspose.app/pdf/conversion/docx-to-base64

# Known issues (front-end bugfixes needed):

After adding a new component (assuming it works and the UI updates correctly)...

1. Need to click "reload" button in the app to update React state list of components in the document (need to fix this to work automatically in the code after the ./src/state/components.js insertTag function)
2. Need to add a new line break after the added component. If the component is the last thing in the document, then it's difficult for the user to add content below it. Manually, user needs to: A) Click inside the new component, at the end of its contnent. B) Press "right arrow" a couple times.

# Instructions (how to make this work in Microsoft Word Online):

1. Open any Word document in Sharepoint or Onedrive. Anything. Create a new one. Can be your personal, or corporate, whatever.
2. Save this manifest.xml file to your computer: https://word-editor-poc.vercel.app/manifest.xml
3. See screenshot. Upload this manifest file to MS Word to install the Add-in. Find the "Add-ins" button in MS Word's toolbar. Then find the option to "Upload My Add-in".
4. Click the "TEMPLATE EDITOR" button in the top toolbar.
   ![image](https://github.com/paulshorey/word-editor-poc/assets/7524065/44eadb91-c688-4e34-a572-3a2821ca5fc2)

# Concerns and problems:

Word Online development is weird and quirky, but overall it works well and makes a beautiful UI inside the document.

EXCEPT the two issues we've been talking about, which are back-end/dev-ops issues with hosting the MS Word app, not the front-end UI...

1. Inserting new base64/ooxml (formatted content) is not reliable -- requires that the document in sharepoint/onedrive be reloaded. Maybe this is possible from the back-end Aspose service that generated it in the first place, but maybe not. We need back-end dev-ops to investigate.
2. Customizing the MS Word toolbar (disable file save-as, and share functionality)

These issues are NOT front-end. They require back-end / dev-ops / hosting research and development.

# Developing the front-end:

This uses Webpack/EsLint/TypeScript/React, your typical web app. No Next JS. The framework is "Microsoft Office JS" which runs this in a sandbox environment, limited, unable to interact with the rest of Microsoft's app, except for some API that it exposes to us, like "Word" or "Document".

Documentation:

- https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins
- https://learn.microsoft.com/en-us/office/dev/add-ins/design/interface-elements
- https://learn.microsoft.com/en-us/office/dev/add-ins/develop/develop-overview#create-an-office-add-in
- https://learn.microsoft.com/en-us/javascript/api/word?view=word-js-preview

But also, TypeScript. If you have good TypeScript tooling (VsCode helps a lot), then you'll be able to "Cmd + Click" on any of the exposed Microsoft methods/properties, and see the definition, and definitions of siblings. It is important to do this in addition to reading the web docs.
