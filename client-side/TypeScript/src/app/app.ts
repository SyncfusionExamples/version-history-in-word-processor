import { enableRipple, isNullOrUndefined } from '@syncfusion/ej2-base';
import { Button } from '@syncfusion/ej2-buttons';
import {
  createSpinner,
  showSpinner,
  hideSpinner,
  DialogUtility,
} from '@syncfusion/ej2-popups';
enableRipple(true);

// tslint:disable-next-line:max-line-length
import {
  DocumentEditorContainer,
  Toolbar,
  ContainerContentChangeEventArgs,
  DocumentEditor,
  ViewChangeEventArgs, CollaborativeEditingHandler, Operation
} from '@syncfusion/ej2-documenteditor';
import { HubConnectionBuilder, HttpTransportType, HubConnectionState } from '@microsoft/signalr';
import { Grid, Page, CommandColumn } from '@syncfusion/ej2-grids';
import {
  ClickEventArgs,
  NodeSelectEventArgs,
  TreeView,
} from '@syncfusion/ej2-navigations';
import { Dialog } from '@syncfusion/ej2-popups';
import { TitleBar } from './title-bar';
import { gridData } from './grid-datasoruce';
import { Toolbar as NavToolbar } from '@syncfusion/ej2-navigations';
import * as Default from './data-default.json';
import * as CharacterFormatting from './data-character-formatting.json';
import * as ParagraphFormatting from './data-paragraph-formatting.json';
import * as Styles from './data-styles.json';
import * as WebLayout from './data-web-layout.json';
// Registering Syncfusion license key
import { registerLicense } from '@syncfusion/ej2-base';

//Register the Syncfusion generated license https://ej2.syncfusion.com/documentation/licensing/license-key-generation
//registerLicense('Replace your generated license key here');

Grid.Inject(Page, CommandColumn);

let url: string = 'http://localhost:62869/api/documenteditor/'
createSpinner({
  // Specify the target for the spinner to show
  target: document.getElementById('main') as HTMLElement,
});
//let downloadButton
let button = new Button({ iconCss: 'e-download icon' });
button.appendTo('#downloadButton');

let staticData: any = [
  {
    id: '01',
    name: 'February 22',
    user: 'Nancy Davolio',
    expanded: true,
    subChild: [
      {
        id: '22-2024-v4',
        name: '2:46 PM',
        user: 'Nancy Davolio',
      },
      {
        id: '22-2024-v3',
        name: '2:45 PM',
        user: 'Nancy Davolio',
      },
      {
        id: '22-2024-v2',
        name: '2:44 PM',
        user: 'Nancy Davolio',
      },
      {
        id: '22-2024-v1',
        name: '2:43 PM',
        user: 'Nancy Davolio',
      },
    ],
  },
];
let contentChanged: boolean = false;

let downloadButton: HTMLButtonElement;

let dialogUtilityObj: Dialog;

let startPage: number = 1;

let containerDialogObj: Dialog = new Dialog({
  width: '90%',
  height: '90%',
  visible: false,
  enableResize: true,
  isModal: true,
  open: onOpen,
  zIndex: 1500,
  position: { X: 'center', Y: 'center' },
});
containerDialogObj.appendTo('#containerDialog');

let editorDialogObj: Dialog = new Dialog({
  width: '90%',
  height: '90%',
  visible: false,
  enableResize: true,
  isModal: true,
  zIndex: 1500,
  position: { X: 'center', Y: 'center' },
});
editorDialogObj.appendTo('#editorDialog');

// Render the TreeView by mapping its fields property with data source properties
let treeObj: TreeView = new TreeView({
  fields: { dataSource: staticData, id: 'id', text: 'name', child: 'subChild' },
  nodeSelected: compareSelected,
  cssClass: 'custom',
  nodeTemplate: '#treeTemplate',
});
treeObj.appendTo('#tree');
let hostUrl: string =
  'https://services.syncfusion.com/js/production/api/documenteditor/';

//Collaborative editing controller url
let serviceUrl = 'http://localhost:62869/';
let collborativeEditingHandler: CollaborativeEditingHandler;
let connectionId: string = "";
let currentRoomName: string = '';

let container: DocumentEditorContainer = new DocumentEditorContainer({
  height: '600px',
  // serviceUrl: hostUrl,
  zIndex: 3000,
  currentUser: 'Guest User',
});
container.serviceUrl = serviceUrl + 'api/documenteditor/';
DocumentEditorContainer.Inject(Toolbar);
container.appendTo('#container');
//Injecting collaborative editing module
DocumentEditor.Inject(CollaborativeEditingHandler);
//Enable collaborative editing in DocumentEditor
container.documentEditor.enableCollaborativeEditing = true;


let documentEditor: DocumentEditor = new DocumentEditor({
  height: '550px',
  serviceUrl: hostUrl,
  zIndex: 3000,
});
// Enable all built in modules.
documentEditor.enableAllModules();
documentEditor.appendTo('#documentEditor');

let toolbarClick: any = function (args: ClickEventArgs) {
  let text = args.item.text;
  switch (text) {
    case 'Edit Document':
      editorDialogObj.hide();
      loadLatestVersion(documentEditor.documentName + '.docx');
      container.documentEditor.enableContextMenu = true;
      container.documentEditor.isReadOnly = false;
      container.showPropertiesPane = true;
      (document.getElementById('documenteditor-share') as HTMLElement).style.display = 'block';
      container.toolbarItems = [
        'New',
        'Open',
        'Separator',
        'Undo',
        'Redo',
        'Separator',
        'Image',
        'Table',
        'Hyperlink',
        'Bookmark',
        'TableOfContents',
        'Separator',
        'Header',
        'Footer',
        'PageSetup',
        'PageNumber',
        'Break',
        'InsertFootnote',
        'InsertEndnote',
        'Separator',
        'Find',
        'Separator',
        'Comments',
        'TrackChanges',
        'Separator',
        'LocalClipboard',
        'RestrictEditing',
        'Separator',
        'FormFields',
        'UpdateFields',
      ];
      containerDialogObj.show();
      break;
    case 'Save a copy':
      downloadDocument();
      break;
    default:
      editorDialogObj.hide();
      break;
  }
};

// Initialize Toolbar component
let toolbarObj: NavToolbar = new NavToolbar({
  clicked: toolbarClick,
  items: [
    {
      prefixIcon: 'e-home icon',
      tooltipText: 'Home',
      text: 'Home',
    },
    {
      prefixIcon: 'e-edit icon',
      tooltipText: 'Edit the latest version',
      text: 'Edit Document',
      align: 'Center',
    },
    {
      prefixIcon: 'e-save icon',
      tooltipText: 'Save a copy',
      text: 'Save a copy',
      align: 'Center',
    },
    {
      prefixIcon: 'e-btn-icon e-icons e-close',
      tooltipText: 'Close version history',
      align: 'Right',
    },
  ],
});

// Render initialized Toolbar
toolbarObj.appendTo('#editorToolbar');
function onOpen(): void {
  //container.height= '94%';
}
let titleBar: TitleBar = new TitleBar(
  document.getElementById('documenteditor_titlebar') as HTMLElement,
  container.documentEditor,
  true,
  false,
  containerDialogObj
);
container.documentChange = (): void => {
  titleBar.updateDocumentTitle();
  container.documentEditor.focusIn();
  titleBar.saveOnClose = false;
};
let operations: any = [];
container.contentChange = (args: ContainerContentChangeEventArgs): void => {
  if (container.documentEditor.enableCollaborativeEditing) {
    //TODO add collaborative editing related code logic when enabling collaborative editing.
    //Send the editing action to server
    collborativeEditingHandler.sendActionToServer(args.operations as Operation[]);
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
//Auto save is triggered based on the timer, we used 15 seconds.
setInterval(() => {
  if (contentChanged) {
    saveDocument();
    contentChanged = false;
  }
}, 15000);

// SignalR connection
var connection = new HubConnectionBuilder().withUrl(serviceUrl + 'documenteditorhub', {
  skipNegotiation: true,
  transport: HttpTransportType.WebSockets
}).withAutomaticReconnect().build();

async function connectToRoom(data: any) {
  try {
    currentRoomName = data.roomName;
    // start the connection.
    connection.start().then(function () {
      // Join the room.
      connection.send('JoinGroup', { roomName: data.roomName, currentUser: data.currentUser });
      console.log('server connected!!!');
    });
  } catch (err) {
    console.log(err);
    //Attempting to reconnect in 5 seconds
    setTimeout(connectToRoom, 5000);
  }
};

connection.onreconnected(() => {
  if (currentRoomName != null) {
    connection.send('JoinGroup', { roomName: currentRoomName, currentUser: container.currentUser });
  }
  console.log('server reconnected!!!');
});

//Event handler for signalR connection
connection.on('dataReceived', onDataRecived.bind(this));


function onDataRecived(action: string, data: any) {
  if (collborativeEditingHandler) {    
    if (action == 'connectionId') {
      //Update the current connection id to track other users
      connectionId = data;
    } else if (connectionId != data.connectionId) {
      if (action == 'action' || action == 'addUser') {
        //Add the user to title bar when user joins the room
        titleBar.addUser(data);
      } else if (action == 'removeUser') {
        //Remove the user from title bar when user leaves the room
        titleBar.removeUser(data);
      }
    }
    //Apply the remote action in DocumentEditor
    collborativeEditingHandler.applyRemoteAction(action, data);
  }
}

connection.onclose(async () => {
  if (connection.state === HubConnectionState.Disconnected) {
    alert('Connection lost. Please relod the browser to continue.');
  }
});

documentEditor.viewChange = (args: ViewChangeEventArgs): void => {
  if (
    documentEditor.selection &&
    documentEditor.selection.startPage >= args.startPage &&
    documentEditor.selection.startPage <= args.endPage
  ) {
    startPage = documentEditor.selection.startPage;
  } else {
    startPage = args.startPage;
  }
  updatePageNumber();
  updatePageCount();
};

function updatePageNumber(): void {
  (document.getElementById('currentPageNumber') as HTMLElement).innerText = startPage.toString();
}
function updatePageCount(): void {
  (document.getElementById('pageCount') as HTMLElement).innerText =
    documentEditor.pageCount.toString();
}

function loadLatestVersion(roomName: string) {
  let responseData: any;
  if (container.documentEditor.enableCollaborativeEditing) {
    responseData = {
      fileName: roomName,
      documentOwner: "fc8094d67084488780c05965ab2f6d53",
    };
  } else {
    responseData = {
      fileName: roomName,
    };
  }
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
        if (container.documentEditor.enableCollaborativeEditing) {
          let data = JSON.parse(httpRequest.responseText);
          collborativeEditingHandler = container.documentEditor.collaborativeEditingHandlerModule;
          //Update the room and version information to collaborative editing handler.
           collborativeEditingHandler.updateRoomInfo(roomName, data.version, serviceUrl+ 'api/documenteditor/');          
          container.documentEditor.open(data.sfdt);
        }
        else {
          container.documentEditor.open(httpRequest.responseText);
        }
        setTimeout(function () {
          // connect to server using signalR
          connectToRoom({ action: 'connect', roomName: roomName, currentUser: container.currentUser });
        });
      }
    }
  };
  httpRequest.send(JSON.stringify(<any>responseData));
}

function compareSelected(args: NodeSelectEventArgs) {
  if (!isNullOrUndefined(args.nodeData.parentID)) {
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
         // documentEditor.open(httpRequest.response);
          let response =JSON.parse(httpRequest.response.sfdt);
          documentEditor.open(response); 
          hideSpinner(document.getElementById('main') as HTMLElement);
        } else {
          hideSpinner(document.getElementById('main') as HTMLElement);
        }
      }
    };
    httpRequest.send(JSON.stringify(<any>responseData));
  } else {
    treeObj.expandAll([args.node]);
  }
}
//  Initialize and render Confirm dialog with options
function showConfirmationDialog(): void {
  dialogUtilityObj = DialogUtility.confirm({
    title: 'Syncfusion Document Editor',
    content: 'Do you want to download a copy of this file and work offline',
    okButton: { text: 'OK', click: okClick },
    cancelButton: { text: 'Cancel', click: cancelClick },
    showCloseIcon: true,
    closeOnEscape: true,
    animationSettings: { effect: 'Zoom' },
  });
}
function okClick(): void {
  downloadDocument();
  //Hide the dialog
  dialogUtilityObj.hide();
}
function cancelClick(): void {
  //Hide the dialog
  dialogUtilityObj.hide();
}

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

function getVersionHistory(name: string) {
  showSpinner(document.getElementById('main') as HTMLElement);
  let responseData: any = {
    fileName: name,
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
function saveDocument() {
  //You can save the document as below
  container.documentEditor.saveAsBlob('Docx').then((blob: Blob) => {
    let fileReader: any = new FileReader();
    fileReader.onload = (): void => {
      let base64String: any = fileReader.result;
      let responseData: any = {
        fileName: container.documentEditor.documentName + '.docx',
        modifiedUser: container.currentUser,
        documentData: base64String,
      };
      let baseUrl: string = url + 'AutoSave';
      let httpRequest: XMLHttpRequest = new XMLHttpRequest();
      httpRequest.open('POST', baseUrl, true);
      httpRequest.setRequestHeader(
        'Content-Type',
        'application/json;charset=UTF-8'
      );
      httpRequest.send(JSON.stringify(<any>responseData));
    };
    fileReader.readAsDataURL(blob);
  });
}
let commandClick: any = function (args: any) {
  let mode = args.target.title;
  let currentDocument = args.rowData.FileName;
  container.documentEditor.documentName = documentEditor.documentName =
    currentDocument.replace('.docx', '');
  titleBar.updateDocumentTitle();
  if (mode !== 'Version History') {
    switch (currentDocument) {
      case 'Getting Started.docx':
        loadLatestVersion(currentDocument);
        break;
      case 'Character Formatting.docx':
        loadLatestVersion(currentDocument);
        break;
      case 'Paragraph Formatting.docx':
        loadLatestVersion(currentDocument);
        break;
      case 'Styles.docx':
        loadLatestVersion(currentDocument);
        break;
      case 'Web layout.docx':
        loadLatestVersion(currentDocument);
        break;
    }
  }
  switch (mode) {
    case 'View':
      container.documentEditor.enableContextMenu = false;
      container.documentEditor.isReadOnly = true;
      container.showPropertiesPane = false;
      (document.getElementById('documenteditor-share') as HTMLElement).style.display = 'none';
      container.toolbarItems = ['Open', 'Separator', 'Find'];
      containerDialogObj.show();
      break;
    case 'Version History':
      getVersionHistory(currentDocument);
      documentEditor.enableContextMenu = false;
      documentEditor.isReadOnly = true;
      documentEditor.showRevisions = false;
      (document.getElementById('documenteditor-share') as HTMLElement).style.display = 'none';
      editorDialogObj.show();
      break;
    default:
      container.documentEditor.enableContextMenu = true;
      container.documentEditor.isReadOnly = false;
      container.showPropertiesPane = true;
      (document.getElementById('documenteditor-share') as HTMLElement).style.display = 'block';
      container.toolbarItems = [
        'New',
        'Open',
        'Separator',
        'Undo',
        'Redo',
        'Separator',
        'Image',
        'Table',
        'Hyperlink',
        'Bookmark',
        'TableOfContents',
        'Separator',
        'Header',
        'Footer',
        'PageSetup',
        'PageNumber',
        'Break',
        'InsertFootnote',
        'InsertEndnote',
        'Separator',
        'Find',
        'Separator',
        'Comments',
        'TrackChanges',
        'Separator',
        'LocalClipboard',
        'RestrictEditing',
        'Separator',
        'FormFields',
        'UpdateFields',
      ];
      containerDialogObj.show();
      break;
  }
};

let grid: Grid = new Grid({
  dataSource: gridData,
  commandClick: commandClick,
  destroyed: destroyed,
  columns: [
    { template: '#fileNameTemplate', headerText: 'File Name' },
    { headerText: 'Author', field: 'Author' },
    {
      textAlign: 'Center',
      headerText: 'Actions',
      commands: [
        { buttonOption: { cssClass: 'e-icons e-eye e-flat' }, title: 'View' },
        { buttonOption: { cssClass: 'e-icons e-edit e-flat' }, title: 'Edit' },
        {
          buttonOption: { cssClass: 'e-icons e-thumbnail e-flat' },
          title: 'Version History',
        },
      ],
    },
  ],
});
grid.appendTo('#Grid');

function destroyed() {
  container.destroy();
  containerDialogObj.destroy();
  editorDialogObj.destroy();
  documentEditor.destroy();
}
