<img width="1594" alt="image" src="https://github.com/paulshorey/word-editor-poc/assets/7524065/0d2c468e-c5d8-4281-a4fc-de4b19ab1c66">

## Resources:

Manually convert Word file to base64 string, then paste into our add-in app text area:
https://products.aspose.app/pdf/conversion/docx-to-base64

## Known issues:

### After adding a new component (assuming it works and the UI updates correctly)...

1. Need to click "reload" button in the app to update React state list of components in the document (need to fix this to work automatically in the code after the ./src/state/components.js insertTag function)
2. Need to add a new line break after the added component. If the component is the last thing in the document, then it's difficult for the user to add content below it. Manually, user needs to: A) Click inside the new component, at the end of its contnent. B) Press "right arrow" a couple times.

## How to make this work in Microsoft Word Online:

1. Open any Word document in Sharepoint or Onedrive. Anything. Create a new one. Can be your personal, or corporate, whatever.
2. Save this manifest.xml file to your computer: base64-word-editor-poc.paulshorey.com/manifest.xml
3. See screenshot. Upload this manifest file to MS Word, to install the Add-in.
4. Click the "TEMPLATE EDITOR" button in the top toolbar.
