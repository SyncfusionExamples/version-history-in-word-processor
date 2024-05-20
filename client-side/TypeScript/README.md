# How to Configure Version History in Client-side.

Please follow the below steps to configure the version history in client-side.

1. Document opened in Document Editor can be saved in the following options.

   a. Save the document based on the number of content changes operation.
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
      saveDocument();
      contentChanged = false;
    } else {
      //Save the document on closing the document irrespective of operations length.
      titleBar.saveOnClose = true;
      contentChanged = false;
    }
     }
   };
   ```
   b. Save the document at periodic time intervals.
    ```
    //Auto save is triggered based on the timer, we used 15 seconds.
    setInterval(() => {
      if (contentChanged) {
        saveDocument();
        contentChanged = false;
      }
    }, 15000);
    ```

    c. Save the current document before closing.
    ```
    private onClose = (): void => {
      if (!this.documentEditor.isReadOnly) {
        this.saveDocument();
      }
      if (this.dialogComponent !== undefined) this.dialogComponent.hide();
    };
    private saveDocument(): void {
      //You can save the document as below
      this.documentEditor.saveAsBlob('Docx').then((blob: Blob) => {
        if (this.saveOnClose) {
          let fileReader: any = new FileReader();
          fileReader.onload = (): void => {
            let base64String: any = fileReader.result;
            let responseData: any = {
              fileName: this.documentEditor.documentName + '.docx',
              modifiedUser: this.documentEditor.currentUser,
              documentData: base64String,
            };

            let baseUrl: string =
              'http://localhost:62869/api/documenteditor/AutoSave';
            let httpRequest: XMLHttpRequest = new XMLHttpRequest();
            httpRequest.open('POST', baseUrl, true);
            httpRequest.setRequestHeader(
              'Content-Type',
              'application/json;charset=UTF-8'
            );
            httpRequest.send(JSON.stringify(<any>responseData));
          };
          fileReader.readAsDataURL(blob);
        }
      });
    }
    ```

2. Saved version history in the server is retrieved for displaying versions on the client-side.

    ```
    function getVersionHistory(name: string) {
      showSpinner(document.getElementById('main') as HTMLElement);
      let responseData: any = {
        DocumentName: name,
      };
      let baseUrl: string = url + 'GetVersionData';
      let httpRequest: XMLHttpRequest = new XMLHttpRequest();
      httpRequest.open('POST', baseUrl, true);
      httpRequest.setRequestHeader(
        'Content-Type',
        'application/json;charset=UTF-8'
      );
      httpRequest.onreadystatechange = () => {
        if (httpRequest.readyState === 4) {
          if (httpRequest.status === 200 || httpRequest.status === 304) {
            let response = JSON.parse(httpRequest.responseText);
            documentEditor.open(response.Document);
            treeObj.fields = {
              dataSource: response.Data,
              id: 'id',
              text: 'name',
              child: 'subChild',
            };
            hideSpinner(document.getElementById('main') as HTMLElement);
          } else {
            hideSpinner(document.getElementById('main') as HTMLElement);
          }
        }
      };
      httpRequest.send(JSON.stringify(<any>responseData));
    }
    ```
3. Retrieve the latest version of the document for editing from the server.
    ```
    function loadLatestVersion(name: string) {
      let responseData: any = {
        DocumentName: name,
      };
      let baseUrl: string = url + 'LoadLatestVersionDocument';
      let httpRequest: XMLHttpRequest = new XMLHttpRequest();
      httpRequest.open('POST', baseUrl, true);
      httpRequest.setRequestHeader(
        'Content-Type',
        'application/json;charset=UTF-8'
      );
      httpRequest.onreadystatechange = () => {
        if (httpRequest.readyState === 4) {
          if (httpRequest.status === 200 || httpRequest.status === 304) {
            container.documentEditor.open(httpRequest.responseText);
          }
        }
      };
      httpRequest.send(JSON.stringify(<any>responseData));
    }
    ```
4. Compare the selected version in the tree view with the previous version.
    ```
    function compareSelected(args: NodeSelectEventArgs) {
      if(!isNullOrUndefined(args.nodeData.parentID)){
      showSpinner(document.getElementById('main') as HTMLElement);
      if (downloadButton) {
        (downloadButton as HTMLButtonElement).style.display = 'none';
      }
      let treeViewRowElement: HTMLElement = args.node.querySelector('.e-text-content') as HTMLElement;
      // Get the button element
      downloadButton = treeViewRowElement.querySelector('.download-button') as HTMLButtonElement;

      // Set the 'display' property to 'none' to hide the button
      (downloadButton as HTMLButtonElement).style.display = 'block';
      downloadButton.addEventListener('click', () => {
          showConfirmationDialog();
      });
      let responseData: any = {
        DocumentName: container.documentEditor.documentName + '.docx',
        SelectedVersion: treeObj.selectedNodes[0],
      };

      let baseUrl: string = url + 'CompareSelectedVersion';
      let httpRequest: XMLHttpRequest = new XMLHttpRequest();
      httpRequest.open('POST', baseUrl, true);
      httpRequest.setRequestHeader(
        'Content-Type',
        'application/json;charset=UTF-8'
      );
      httpRequest.responseType = 'json';
      httpRequest.onreadystatechange = () => {
        if (httpRequest.readyState === 4) {
          if (httpRequest.status === 200 || httpRequest.status === 304) {
            documentEditor.open(httpRequest.response);
            hideSpinner(document.getElementById('main') as HTMLElement);
          } else {
            hideSpinner(document.getElementById('main') as HTMLElement);
          }
        }
      };
      httpRequest.send(JSON.stringify(<any>responseData));
    }else{
      treeObj.expandAll([args.node]);
    }
    }
    ```
5. Download the selected version as word document.
    ```
    function downloadDocument() {
      let documentVersion = treeObj.selectedNodes[0];
      if (isNullOrUndefined(documentVersion)) {
        documentVersion = treeObj.getTreeData()[0].id as string;
      }
      let responseData: any = {
        DocumentName: container.documentEditor.documentName + '.docx',
        SelectedVersion: documentVersion,
      };
      let baseUrl: string = url + 'Download';
      let httpRequest: XMLHttpRequest = new XMLHttpRequest();
      httpRequest.open('POST', baseUrl, true);
      httpRequest.setRequestHeader(
        'Content-Type',
        'application/json;charset=UTF-8'
      );
      httpRequest.responseType = 'blob';
      // Set up event listener for the response
      httpRequest.onload = function () {
        if (httpRequest.status === 200) {
          // Handle the response blob here
          let responseData = httpRequest.response;
          // Create a Blob URL for the response data
          let blobUrl = URL.createObjectURL(responseData);
          // Create a link element and trigger the download
          let downloadLink = document.createElement('a');
          downloadLink.href = blobUrl;
          downloadLink.download =
            container.documentEditor.documentName + '_' + treeObj.selectedNodes[0];
          document.body.appendChild(downloadLink);
          downloadLink.click();
          // Cleanup: Remove the link and revoke the Blob URL
          document.body.removeChild(downloadLink);
          URL.revokeObjectURL(blobUrl);
        } else {
          // Handle errors
          console.error('Request failed with status:', httpRequest.status);
        }
      };
      httpRequest.send(JSON.stringify(<any>responseData));
    }
    ```
