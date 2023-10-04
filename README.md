This is the list of items we want to address. It is the original list made by Paul
with a few additional item.
Each item haas a number in parentheses, that is Karsten's priorities.
When an item has been completed, update the list adding the line with `\*\*`

When you opick an item to work on - please replace the '-' at the first line with your first letter of first name

[X] (9) Show/Hide side column “taskpane” (what Microsoft calls it)

[ ] (4) From the taskpane, insert a “variable” (Data Element) into the document.

[ ] (10) In the document, click the variable to highlight it and show it in the taskpane.

[ ] (5) From the taskpane, clicking the variable should highlight it in the document and scroll to it.

[ ] (6) From the taskpane, edit or delete the variable.

[ ] (#) In the document, when typing “{SOMETHING}” surrounded by curly brackets, assume
that the user is trying to insert a variable. We’ll need to monitor what the user types,
and detect this sort of thing.

[ ] (#) If it is valid, then go ahead and insert it into the state of the taskpane app. If
it’s not valid, then highlight that variable name including curly brackets, in red (to
show that it’s invalid).

[ ] (7) When the document loads, programmatically run through all the variables in the document,
and load them into the taskpane (app) data state. >>> I think the variable definitions are in the document when it is a Template.
When we start editor with on a Template Instance, I think we get the values for the variable.
I think if we repeat starting editing an instance that we keep inserting the variable values

[ ] (1) Insert a “Component” — don’t bother trying to load a whole other Word Document yet. Just
make a distinction in the app between a “Data Element” variable and a “Component” variable. Have
UI to differentiate, and choose which is a component, and which is a data-element. Maybe to make
it super simple at first, just prepend “component\_” to the start of the variable name, and just
display it like that in the document, like “{component_SOME_UNIQUE_ID}”. This is how we will
save the “.docx” file — we can’t save the rendered content. We’ll get that later.

[ ] (2) Now, lets try to actually display the contents of the component variable into the document.
So now, instead of displaying it just like the name, like “{component_something}”, actually print
out the value of it (some HTML or a paragraph of text at least). BUT, the whole entire thing still
needs to be clickable. On click, select the whole thing, and draw an outline around it, and
bring it up in the taskpane, just like with the data-element variable.

[ ] (3A) When the document loads/reloads, we must now fetch the contents for each of these “component”
variables, and add it to our app state, and also update the document to display the content as well.

[ ] (8) Now, lets add ”active”/”hidden”/”flattened” status functionality. On document load,
when fetching the content for each component, also fetch the status. Sure, these statuses are not
what we currently use. They are simplified for the POC. - “active” - display it as usual - editable, clickable, selectable - “hidden” - do not display it at all - “flattened” - display the content, but not as a component, not interactive, but inline as normal content

- (3B) When we save the document we must pull out the contents of the components and just leave a context control
  that contains enough information for extracting the component at next load.

- () When need to be able to listen to events inside word. More specifically, if the user selects a
  range of text that contains either a Variable or a Component, we need to know so the side panel
  can be updated accordingly.
  We could also create variables and component as content controls and make them "not deletable", whichs
  means they can only be deleted from the side panel.
