# How Syncfusion Document Editor Works
Syncfusion Document Editor leverages the power of JavaScript to deliver a feature-rich document editing experience. Here's how it works:
1.	**Integration:** Simply integrate Syncfusion Document Editor into your web application using JavaScript. With easy-to-follow documentation and sample code snippets, getting started is quick and hassle-free.
2.	**Document Editing:** Create, edit, and format your documents using Syncfusion Document Editor's intuitive interface. Whether you're writing a report, crafting a proposal, or designing a presentation, Syncfusion Document Editor offers a comprehensive set of editing tools to help you bring your ideas to life.
3.	**Versioning and Saving:** As you work on your document, Syncfusion Document Editor automatically saves your progress at regular intervals, ensuring that your changes are always preserved. Additionally, you can manually save your document or trigger saves programmatically to suit your workflow.
```
container.contentChange = (args: ContainerContentChangeEventArgs): void => {
  if (container.documentEditor.enableCollaborativeEditing) {
    //TODO add collaborative editing related code logic when enabling collaborative editing.
  } else {
    operations.push(args.operations);
    //Populate the operation upto 50 and auto save the version.
    if (operations.length > 50) {
      contentChanged = true;
      titleBar.saveOnClose = false;
      operations = [];
      autoSaveDocument();
      contentChanged = false;
    } else {
      //Save the document on closing the document irrespective of operations length.
      titleBar.saveOnClose = true;
      contentChanged = false;
    }
  }
};
//Auto save is triggered based on the timer, we used 15 seconds.
setInterval(() => {
  if (contentChanged) {
    autoSaveDocument();
    contentChanged = false;
  }
}, 15000);
```
4.	**Version Comparison:** When you need to review changes or compare document versions, simply access Syncfusion Document Editor's version comparison feature. Instantly visualize the differences between versions, track modifications, and collaborate effectively with your team members.
