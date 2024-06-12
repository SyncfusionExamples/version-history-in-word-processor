import { createRoot } from 'react-dom/client';
import './index.css';
import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { DialogComponent } from '@syncfusion/ej2-react-popups';
import { GridComponent, ColumnsDirective, ColumnDirective, Inject, CommandColumn } from '@syncfusion/ej2-react-grids';
import { TreeViewComponent, ToolbarComponent, ItemsDirective, ItemDirective } from '@syncfusion/ej2-react-navigations';
import { gridData } from './word-data';
import { DocumentEditorContainerComponent, Toolbar, Operation, CollaborativeEditingHandler, ContainerContentChangeEventArgs } from '@syncfusion/ej2-react-documenteditor';
import { DocumentEditor } from '@syncfusion/ej2-react-documenteditor';
import { TitleBar } from './title-bar';
import { HubConnectionBuilder, HttpTransportType, HubConnectionState, HubConnection } from '@microsoft/signalr';

DocumentEditorContainerComponent.Inject(Toolbar);
import { defaultDocument, characterFormat, paragraphFormat, styles, weblayout } from './data';
const DocumentList = () => {
  useEffect(() => {
    rendereComplete();
  }, []);
  const [dictionary, setDictionary] = useState({
    'Getting Started.docx': defaultDocument,
    'Character Formatting.docx': characterFormat,
    'Paragraph Format.docx': paragraphFormat,
    'Style.docx': styles,
    'Web Layout.docx': weblayout
  });
  const commands = [
    { type: 'View', buttonOption: { cssClass: "e-icons e-eye e-flat" } },
    { type: 'Edit', buttonOption: { cssClass: "e-icons e-edit e-flat" } },
    { type: 'View', buttonOption: { cssClass: "e-icons e-thumbnail e-flat" } },
  ];
  let dialogInstance = useRef(null);
  let editorDialogInstance = useRef(null);
  const gridInstance = useRef(null);
  const [isDialogOpen, setDialogOpen] = useState(false);
  const [isEditorDialogOpen, setEditorDialogOpen] = useState(false);
  const [isGridOpen, setGridOpen] = useState(true);
  const [visible, setVisible] = useState(false);
  const [name, setName] = useState('');
  let hostUrl = "https://services.syncfusion.com/react/production/api/documenteditor/";
  let container = useRef(null);
  let editorcontainer = useRef(null);
  let treeObj = useRef(null);
  let titleBar;
  let serviceUrl = 'http://localhost:62869/';
  let collaborativeEditingHandler ='';
  let connectionId = "";
  let currentRoomName = '';
  let currentUser = 'Guest user';
  let operations = [];
  let contentChanged = false;
  let connection;
 
  let selectedNode;
  let url = 'http://localhost:62869/api/documenteditor/';
  let staticData = [
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
  let fields = { dataSource: staticData, id: 'id', text: 'name', child: 'subChild' };  
  let downloadButton;
  const onLoadDefault = () => {
    titleBar.updateDocumentTitle();
    container.current.documentChange = () => {
      titleBar.updateDocumentTitle();
      container.current.documentEditor.focusIn();
    };
    container.current.documentEditorSettings.showRuler = true;
  };
  const rendereComplete = () => {
    window.onbeforeunload = function () {
      return "Want to save your changes?";
    };
    container.current.documentEditor.pageOutline = "#E0E0E0";
    container.current.documentEditor.acceptTab = true;
    container.current.documentEditor.resize();
    titleBar = new TitleBar(document.getElementById("documenteditor_titlebar"), container.current.documentEditor, true, false, dialogInstance.current);
     //Inject the collaborative editing handler to DocumentEditor
     DocumentEditor.Inject(CollaborativeEditingHandler);
     //Enable the collaborative editing in DocumentEditor
     container.current.documentEditor.enableCollaborativeEditing = true;
    onLoadDefault();
    collaborativeEditingHandler = container.current.documentEditor.collaborativeEditingHandlerModule;
    if (container.current.documentEditor.enableCollaborativeEditing) {
    initializeSignalR();
    }
    container.current.contentChange = (args) => {
      if (container.current.documentEditor.enableCollaborativeEditing) {      
        //Send the editing action to server
        collaborativeEditingHandler.sendActionToServer(args.operations);
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
  };
//Collaborative editing 
  const initializeSignalR = () => {
    // SignalR connection
    connection = new HubConnectionBuilder().withUrl(serviceUrl + 'documenteditorhub', {
        skipNegotiation: true,
        transport: HttpTransportType.WebSockets
    }).withAutomaticReconnect().build();
    //Event handler for signalR connection
    connection.on('dataReceived', onDataRecived);

    connection.onclose(async () => {
        if (connection && connection.state === HubConnectionState.Disconnected) {
            alert('Connection lost. Please relod the browser to continue.');
        }
    });
    connection.onreconnected(() => {
        if (connection && currentRoomName != null) {
            connection.send('JoinGroup', { roomName: currentRoomName, currentUser: currentUser });
        }
        console.log('server reconnected!!!');
    });
};

 const onDataRecived = (action, data) => {
    if (collaborativeEditingHandler) {
        if (action == 'connectionId') {
            //Update the current connection id to track other users
            connectionId = data;
        } else if (connectionId != data.connectionId) {
            if (titleBar) {
                if (action == 'action' || action == 'addUser') {
                    //Add the user to title bar when user joins the room
                   // titleBar.addUser(data);
                } else if (action == 'removeUser') {
                    //Remove the user from title bar when user leaves the room
                   // titleBar.removeUser(data);
                }
            }
        }
        //Apply the remote action in DocumentEditor
        collaborativeEditingHandler.applyRemoteAction(action, data);
    }
};
const connectToRoom = (data) => {
  try {
      currentRoomName = data.roomName;
      if (connection) {
          // start the connection.
          connection.start().then(() => {
              // Join the room.
              if (connection) {
                  connection.send('JoinGroup', { roomName: data.roomName, currentUser: data.currentUser });
              }
              console.log('server connected!!!');
          });
      }

  } catch (err) {
      console.log(err);
      //Attempting to reconnect in 5 seconds
      setTimeout(connectToRoom, 5000);
  }
};
  const saveDocument = () => {
    //You can save the document as below
    container.current.documentEditor.saveAsBlob('Docx').then((blob) => {
      let fileReader = new FileReader();
      fileReader.onload = () => {
        let base64String = fileReader.result;
        let responseData = {
          fileName: container.current.documentEditor.documentName + '.docx',
          modifiedUser: currentUser,
          documentData: base64String,
        };
        let baseUrl = url + 'AutoSave';
        let httpRequest = new XMLHttpRequest();
        httpRequest.open('POST', baseUrl, true);
        httpRequest.setRequestHeader(
          'Content-Type',
          'application/json;charset=UTF-8'
        );
        httpRequest.send(JSON.stringify(responseData));
      };
      fileReader.readAsDataURL(blob);
    });
  };

  const dialogClose = () => {
    setDialogOpen(false);
  };
  const dialogOpen = () => {
    setDialogOpen(true);
    container.current.documentEditor.resize();
  };
  const editorDialogOpen = () => {
    setEditorDialogOpen(true);
    container.current.documentEditor.resize();
  };

  const editorDialogClose = () => {
    setEditorDialogOpen(false);
  };
  const getVersionHistory = (name) => {
    // showSpinner(document.getElementById('main'));
    let responseData = {
      fileName: name,
    };
    let baseUrl = url + 'GetVersionData';
    let httpRequest = new XMLHttpRequest();
    httpRequest.open('POST', baseUrl, true);
    httpRequest.setRequestHeader(
      'Content-Type',
      'application/json;charset=UTF-8'
    );
    httpRequest.onreadystatechange = () => {
      if (httpRequest.readyState === 4) {
        if (httpRequest.status === 200 || httpRequest.status === 304) {
          let response = JSON.parse(httpRequest.responseText);
          editorcontainer.current.documentEditor.open((response.Document));

          treeObj.current.fields = {
            dataSource: response.Data,
            id: 'id',
            text: 'name',
           //text : setName('name'),
            child: 'subChild',
          };
          
        } 
      }
    };
    httpRequest.send(JSON.stringify(responseData));
  };

  const loadLatestVersion = (roomName) => {
    let responseData;
    if (container.current.documentEditor.enableCollaborativeEditing) {
      responseData = {
        fileName: roomName,
        documentOwner: "fc8094d67084488780c05965ab2f6d53",
      };
    } else {
      responseData = {
        fileName: roomName,
      };
    }
    let baseUrl = url + 'LoadLatestVersionDocument';
    let httpRequest = new XMLHttpRequest();
    httpRequest.open('POST', baseUrl, true);
    httpRequest.setRequestHeader(
      'Content-Type',
      'application/json;charset=UTF-8'
    );
    httpRequest.onreadystatechange = () => {
      if (httpRequest.readyState === 4) {
        if (httpRequest.status === 200 || httpRequest.status === 304) {
          if (container.current.documentEditor.enableCollaborativeEditing) {
            let data = JSON.parse(httpRequest.responseText);
            collaborativeEditingHandler = container.current.documentEditor.collaborativeEditingHandlerModule;
            //Update the room and version information to collaborative editing handler.
            collaborativeEditingHandler.updateRoomInfo(roomName, data.version, serviceUrl + 'api/documenteditor/');
            container.current.documentEditor.open(data.sfdt);
            setTimeout(function () {
               connectToRoom({ action: 'connect', roomName: roomName, currentUser: currentUser });
            });
          }
          else {
            container.current.documentEditor.open(httpRequest.responseText);
          }
         
        }
      }
    };
    httpRequest.send(JSON.stringify(responseData));
  };
  const onCommandClicked = (args) => {
    const cssClass = args.target.className;
    const currentDocument = args.rowData.FileName;
    container.current.documentEditor.documentName = currentDocument.replace('.docx', '');
    editorcontainer.current.documentEditor.documentName = currentDocument.replace('.docx', '');
    if (cssClass.includes('e-icons e-eye e-flat')) {
      setDialogOpen(true);
      setGridOpen(false);
      if (dictionary.hasOwnProperty(args.rowData.FileName)) {
        switch (currentDocument) {
          case 'Getting Started.docx':
            loadLatestVersion(container.current.documentEditor.documentName);
            break;
          case 'Character Formatting.docx':
            loadLatestVersion(container.current.documentEditor.documentName);
            break;
          case 'Paragraph Formatting.docx':
            loadLatestVersion(container.current.documentEditor.documentName);
            break;
          case 'Styles.docx':
            loadLatestVersion(container.current.documentEditor.documentName);
            break;
          case 'Web layout.docx':
            loadLatestVersion(container.current.documentEditor.documentName);
            break;
        }
      }
      container.current.documentEditor.isReadOnly = true;
      container.current.documentEditor.enableContextMenu = false;
      container.current.resize();
      const downloadButton = document.getElementById("documenteditor-share");
      if (downloadButton) {
        downloadButton.style.display = "none";
      }
      const closeButton = document.getElementById("de-close");
      if (closeButton) {
        closeButton.style.display = "block";
      }
      container.current.documentEditor.documentName = args.rowData.FileName.replace(".docx", "");
      document.getElementById("documenteditor_title_name").textContent = container.current.documentEditor.documentName;
      container.current.toolbarItems = ['Open', 'Separator', 'Find'];
    }
    else if (cssClass.includes('e-icons e-edit e-flat')) {
      setDialogOpen(true);
      setGridOpen(false);
      if (dictionary.hasOwnProperty(args.rowData.FileName)) {
        switch (currentDocument) {
          case 'Getting Started.docx':
            loadLatestVersion(container.current.documentEditor.documentName);
            break;
          case 'Character Formatting.docx':
            loadLatestVersion(container.current.documentEditor.documentName);
            break;
          case 'Paragraph Formatting.docx':
            loadLatestVersion(container.current.documentEditor.documentName);
            break;
          case 'Styles.docx':
            loadLatestVersion(container.current.documentEditor.documentName);
            break;
          case 'Web layout.docx':
            loadLatestVersion(container.current.documentEditor.documentName);
            break;
        }
      }
      container.current.documentEditor.isReadOnly = false;
      container.current.documentEditor.enableContextMenu = true;
      container.current.resize();
      const downloadButton = document.getElementById("documenteditor-share");
      if (downloadButton) {
        downloadButton.style.display = "block";
      }
      const closeButton = document.getElementById("de-close");
      if (closeButton) {
        closeButton.style.display = "block";
      }
      container.current.documentEditor.documentName = args.rowData.FileName.replace(".docx", "");
      document.getElementById("documenteditor_title_name").textContent = container.current.documentEditor.documentName;
      container.current.toolbarItems = ['New', 'Open', 'Separator', 'Undo', 'Redo', 'Separator', 'Image', 'Table', 'Hyperlink', 'Bookmark', 'TableOfContents', 'Separator', 'Header', 'Footer', 'PageSetup', 'PageNumber', 'Break', 'InsertFootnote', 'InsertEndnote', 'Separator', 'Find', 'Separator', 'Comments', 'TrackChanges', 'Separator', 'LocalClipboard', 'RestrictEditing', 'Separator', 'FormFields', 'UpdateFields'];
    } else if (cssClass.includes('e-icons e-thumbnail e-flat')) {
      getVersionHistory(args.rowData.FileName);
      editorcontainer.current.documentEditor.enableContextMenu = false;
      editorcontainer.current.documentEditor.isReadOnly = true;
      editorcontainer.current.documentEditor.showRevisions = false;
      const downloadButton = document.getElementById("documenteditor-share");
      if (downloadButton) {
        downloadButton.style.display = "none";
      }
      setEditorDialogOpen(true);
    }
  };
  function compareSelected(args) {   
    if (downloadButton) {
      (downloadButton).style.display = 'none';
    }
    if ((args.nodeData.parentID) != null) {     
      let treeViewRowElement = args.node.querySelector('.e-text-content');      
      downloadButton = treeViewRowElement.querySelector('.e-de-icon-Download');  
      if (downloadButton) {
        downloadButton.style.display = 'block';

        downloadButton.addEventListener('click', () => {         
          showConfirmationDialog();
        });
      }
      let responseData = {
        DocumentName: editorcontainer.current.documentEditor.documentName + '.docx',       
        SelectedVersion: treeObj.current.selectedNodes[0],
      };

      let baseUrl = url + 'CompareSelectedVersion';
      let httpRequest = new XMLHttpRequest();
      httpRequest.open('POST', baseUrl, true);
      httpRequest.setRequestHeader(
        'Content-Type',
        'application/json;charset=UTF-8'
      );     
      httpRequest.onreadystatechange = () => {
        if (httpRequest.readyState === 4) {
          if (httpRequest.status === 200 || httpRequest.status === 304) {

            let response = JSON.parse(httpRequest.response);
            editorcontainer.current.documentEditor.open((response.sfdt));          
          } 
        }
      };
      httpRequest.send(JSON.stringify(responseData));
    } else {
      treeObj.current.expandAll([args.node]);
    }
  }
  //  Initialize and render Confirm dialog with options
  function showConfirmationDialog() {
    setVisible(true);
  }

  function okClick() {
    downloadDocument();    
    setVisible(false);
    
  }
  function cancelClick() { 
    setVisible(false);   
  }

  const downloadDocument = () => {    
    let documentVersion = treeObj.current.selectedNodes[0];
    if (documentVersion == null) {
      documentVersion = treeObj.current.getTreeData()[0].id;
    }  
    let responseData = {
      DocumentName: editorcontainer.current.documentEditor.documentName + '.docx',
      SelectedVersion: documentVersion,
    };
    let baseUrl = url + 'Download';
    let httpRequest = new XMLHttpRequest();
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
          editorcontainer.current.documentEditor.documentName + '_' + treeObj.current.selectedNodes[0];
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
   
    httpRequest.send(JSON.stringify(responseData));
  }

  const toolbarClick = (args) => {
    console.log(args.item);
    let text = args.item.text;
    switch (text) {
      case 'Edit Document':
        setEditorDialogOpen(false);
        loadLatestVersion(container.current.documentEditor.documentName + '.docx');
        container.current.documentEditor.enableContextMenu = true;
        container.current.documentEditor.isReadOnly = false;
        container.showPropertiesPane = true;
        const downloadButton = document.getElementById("documenteditor-share");
        if (downloadButton) {
          downloadButton.style.display = "block";
        }
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
        setDialogOpen(true);
        break;
      case 'Save a copy':
        downloadDocument();
        break;
      default:
        setEditorDialogOpen(false);
        break;
    }
  };
  function nodeTemplate(data) {
    return (<div>
     <div style={{ display: 'flex', alignitems: 'center' }}>
          <div className="ename" style={{ marginright: '30px' }}>{data.name}</div>
          <button id="downloadButton" className="e-btn-icon e-de-icon-Download e-de-padding-right e-icon-left" title="Download a copy of the document." data-ripple="true"></button>
        </div>
        <svg width="20" height="20" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
          <rect x="4" y="4" width="12" height="12" rx="2" fill="gray" />
        </svg>
        <span style={{ verticalAlign: 'super', marginRight: '10px' }}>{data.user}</span>
        <span style={{ verticalAlign: 'super' }}>modified</span>


  </div>);
}
  return (
    <div className="control-pane documenteditor-list-sample">
      <GridComponent ref={gridInstance} dataSource={gridData} commandClick={onCommandClicked}>
        <ColumnsDirective>
          <ColumnDirective headerText='File Name' template={(props) => (<div className="file-name-container">
            <div className="file-name-content">
              <div className="icon-and-text">
                <svg width="30" height="30" viewBox="0 0 30 30" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M3 3C3 1.34315 4.34315 0 6 0H16.7574C17.553 0 18.3161 0.316071 18.8787 0.87868L26.1213 8.12132C26.6839 8.68393 27 9.44699 27 10.2426V27C27 28.6569 25.6569 30 24 30H6C4.34315 30 3 28.6569 3 27V3Z" fill="#4889EF" />
                  <path d="M17.5 11H25V10.5042C25 9.76949 24.7304 9.0603 24.2422 8.51114L19.9463 3.67818C18.9974 2.61074 17.6374 2 16.2092 2H16V9.5C16 10.3284 16.6716 11 17.5 11Z" fill="#D6E5FE" />
                  <path d="M10.3044 12H10.8868H11.104H11.6817L12.6231 16.3922L13.3963 12H15L13.5719 19H12.777H12.5552H11.8943L10.993 15.0093L10.1103 19H9.44945H9.22761H8.42808L7 12H8.60832L9.38188 16.3816L10.3044 12Z" fill="white" />
                  <rect x="7" y="21" width="16" height="2" rx="1" fill="white" />
                  <rect x="7" y="25" width="11" height="2" rx="1" fill="white" />
                </svg>
                <div className="file-name-text">{props.FileName}</div>
              </div>
            </div>
          </div>)} />
          <ColumnDirective headerText='Author' field='Author'></ColumnDirective>
          <ColumnDirective headerText='Actions' commands={commands} textAlign='Center'></ColumnDirective>
        </ColumnsDirective>
        <Inject services={[CommandColumn]} />
      </GridComponent>
      <DialogComponent id="defaultDialog" ref={dialogInstance} isModal={true} visible={isDialogOpen} width={'90%'} height={'90%'} zIndex={1500} open={dialogOpen} close={dialogClose} minHeight={'650px'}>
        <div>
          <div id="documenteditor_titlebar" className="e-de-ctn-title"></div>
          <div id="documenteditor_container_body">
            <DocumentEditorContainerComponent showPropertiesPane={false} id="container" height='780px' ref={container} style={{ display: "block" }} serviceUrl={hostUrl} zIndex={3000} enableToolbar={true} locale="en-US" />
          </div>
        </div>
      </DialogComponent>
      <DialogComponent id="editorDialog" ref={editorDialogInstance} isModal={true} visible={isEditorDialogOpen} width={'90%'} height={'90%'} zIndex={1500} open={editorDialogOpen} close={editorDialogClose} minHeight={'650px'}>
        <div id="editorToolbar">  
        <ToolbarComponent clicked={toolbarClick}>
          <ItemsDirective>
            <ItemDirective
              prefixIcon='e-home icon'
              tooltipText='Home'
              text='Home'
            />
            <ItemDirective
              prefixIcon='e-edit icon'
              tooltipText='Edit the latest version'
              text='Edit Document'
              align='Center'
            />
            <ItemDirective
              prefixIcon='e-save icon'
              tooltipText='Save a copy'
              text='Save a copy'
              align='Center'
            />
            <ItemDirective
              prefixIcon='e-btn-icon e-icons e-close'
              tooltipText='Close version history'
              align='Right'
            />
          </ItemsDirective>
        </ToolbarComponent></div>
        <div id="main" style={{ display: 'flex', width: '100%' }}>
          <div id="sub1" style={{ width: '70%' }}>
            <div>
              <div id="documenteditor_titlebar1" className="e-de-ctn-title"></div>
              <div id="documenteditor_container_body1">
                <DocumentEditorContainerComponent showPropertiesPane={false} id="editorcontainer" height='780px' ref={editorcontainer} style={{ display: "block" }} serviceUrl={hostUrl} zIndex={3000} enableToolbar={false} locale="en-US" />
              </div>
            </div>            
          </div>
          <div id="sub2" style={{ width: '30%' }}>
            <div style={{ fontSize: '24px', marginLeft: '10px' }}>Version History</div>
             <TreeViewComponent id="treeObj" ref={treeObj} fields={fields} nodeSelected={compareSelected.bind(this)} cssClass='custom' nodeTemplate={nodeTemplate} />
            </div>
        </div>
      </DialogComponent>
      <div>
        <DialogComponent
          minHeight={'160px'}
          visible={visible}
          isModal={true}
          header='Syncfusion Document Editor'
          content='Do you want to download a copy of this file and work offline?'
          showCloseIcon={true}
          closeOnEscape={true}
          animationSettings={{ effect: 'Zoom' }}
          buttons={[
            { click: okClick, buttonModel: { content: 'OK', isPrimary: true } },
            { click: cancelClick, buttonModel: { content: 'Cancel' } }
          ]}
          width='400px'
          close={() => setVisible(false)}
        />
      </div>          
      <script id="fileNameTemplate">
        <div>
          <svg width="30" height="30" viewBox="0 0 30 30" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M3 3C3 1.34315 4.34315 0 6 0H16.7574C17.553 0 18.3161 0.316071 18.8787 0.87868L26.1213 8.12132C26.6839 8.68393 27 9.44699 27 10.2426V27C27 28.6569 25.6569 30 24 30H6C4.34315 30 3 28.6569 3 27V3Z" fill="#4889EF" />
            <path d="M17.5 11H25V10.5042C25 9.76949 24.7304 9.0603 24.2422 8.51114L19.9463 3.67818C18.9974 2.61074 17.6374 2 16.2092 2H16V9.5C16 10.3284 16.6716 11 17.5 11Z" fill="#D6E5FE" />
            <path d="M10.3044 12H10.8868H11.104H11.6817L12.6231 16.3922L13.3963 12H15L13.5719 19H12.777H12.5552H11.8943L10.993 15.0093L10.1103 19H9.44945H9.22761H8.42808L7 12H8.60832L9.38188 16.3816L10.3044 12Z" fill="white" />
            <rect x="7" y="21" width="16" height="2" rx="1" fill="white" />
            <rect x="7" y="25" width="11" height="2" rx="1" fill="white" />
          </svg><span style={{ verticalalign: 'super' }}>FileName</span>
        </div>
      </script>
    </div>);
};
export default DocumentList;

const root = createRoot(document.getElementById('sample'));
root.render(<DocumentList />);