This is the list of items we want to address. It is the original list made by Paul
with a few additional item.
Each item haas a number in parentheses, that is Karsten's priorities.
When an item has been completed, update the list adding the line with `\*\*`

When you pick an item to work on - please replace the '-' at the first line with your first letter of first name

## Todo for the first week POC:

[X] (1) Show/Hide side column “taskpane” (what Microsoft calls it)

[X] (2) From the taskpane, insert a “variable” (Data Element) into the document.

[X] (3) In the document, click the variable to highlight it and show it in the taskpane.

[X] (4) From the taskpane, clicking the variable should highlight it in the document and scroll to it.

[X] (5) From the taskpane, delete the variable.

[X] (6) When the document loads, programmatically run through all the variables in the document and load them into the taskpane (app) data state.

[X] (7) Insert a “Component” which may somehow contain contents of another Word document, or at least text/HTML.

[X] (8) Now, lets try to actually display the contents of the component variable into the document.
So now, instead of displaying it just like the name, like “{component_something}”, actually print
out the value of it (some HTML or a paragraph of text at least). BUT, the whole entire thing still
needs to be clickable. On click, select the whole thing, and draw an outline around it, and
bring it up in the taskpane, just like with the data-element variable.

[X] (9) When the document loads/reloads, we must now fetch the contents for each of these “component”
variables, and add it to our app state, and also update the document to display the content as well.

[X] (10) When we save the document we must pull out the contents of the components and just leave a context control
that contains enough information for extracting the component at next load.

[X] (11) When need to be able to listen to events inside word. More specifically, if the user selects a
range of text that contains either a Variable or a Component, we need to know so the side panel
can be updated accordingly.
We could also create variables and component as content controls and make them "not deletable", whichs
means they can only be deleted from the side panel.

[X] (12) Conditional elements

[X] (13) Add state for components - similar to what we have for variables

## Resources:

Link to online site that converts word files to base64 strings
https://products.aspose.app/pdf/conversion/docx-to-base64
