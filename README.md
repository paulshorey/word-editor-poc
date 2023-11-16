<img width="1594" alt="image" src="https://github.com/paulshorey/word-editor-poc/assets/7524065/0d2c468e-c5d8-4281-a4fc-de4b19ab1c66">

## Resources:

Link to online site that converts word files to base64 strings
https://products.aspose.app/pdf/conversion/docx-to-base64

## Known issues:

### After adding a new component (assuming it works and the UI updates correctly)...

1. Need to click "reload" button in the app to update React state list of components in the document (need to fix this to work automatically in the code after the ./src/state/components.js insertTag function)
2. Need to add a new line break after the added component. If the component is the last thing in the document, then it's difficult for the user to add content below it. Manually, user needs to: A) Click inside the new component, at the end of its contnent. B) Press "right arrow" a couple times.
