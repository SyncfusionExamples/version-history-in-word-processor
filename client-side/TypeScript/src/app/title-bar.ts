import { createElement, Event, KeyboardEventArgs } from '@syncfusion/ej2-base';
import { ActionInfo,DocumentEditor, FormatType } from '@syncfusion/ej2-documenteditor';
import { Button } from '@syncfusion/ej2-buttons';
import { DropDownButton, ItemModel } from '@syncfusion/ej2-splitbuttons';
import { MenuEventArgs } from '@syncfusion/ej2-navigations';
import { Dialog } from '@syncfusion/ej2-popups';
/**
 * Represents document editor title bar.
 */
export class TitleBar {
  private tileBarDiv: HTMLElement;
  public saveOnClose: boolean = false;
  private documentTitle?: HTMLElement;
  private documentTitleContentEditor?: HTMLElement;
  private export?: DropDownButton;
  private print?: Button;
  private close?: Button;
  private open?: Button;
  private documentEditor: DocumentEditor;
  private isRtl?: boolean;
  private userList?: HTMLElement;
  public userMap: any = {};
  private dialogComponent?: Dialog;
  constructor(
    element: HTMLElement,
    docEditor: DocumentEditor,
    isShareNeeded: Boolean,
    isRtl?: boolean,
    dialogComponent?: Dialog
  ) {
    //initializes title bar elements.
    this.tileBarDiv = element;
    this.documentEditor = docEditor;
    this.isRtl = isRtl;
    this.dialogComponent = dialogComponent;
    this.initializeTitleBar(isShareNeeded);
    this.wireEvents();
  }
  private initializeTitleBar = (isShareNeeded: Boolean): void => {
    let downloadText: string = '';
    let downloadToolTip: string = '';
    let printText: string = '';
    let printToolTip: string = '';
    let closeToolTip: string = '';
    let openText: string = '';
    let documentTileText: string = '';
    if (!this.isRtl) {
      downloadText = 'Download';
      downloadToolTip = 'Download this document.';
      printText = 'Print';
      printToolTip = 'Print this document (Ctrl+P).';
      closeToolTip = 'Close this document';
      openText = 'Open';
      documentTileText = 'Document Name. Click or tap to rename this document.';
    } else {
      downloadText = 'تحميل';
      downloadToolTip = 'تحميل هذا المستند';
      printText = 'طباعه';
      printToolTip = 'طباعه هذا المستند (Ctrl + P)';
      openText = 'فتح';
      documentTileText = 'اسم المستند. انقر أو اضغط لأعاده تسميه هذا المستند';
    }
    // tslint:disable-next-line:max-line-length
    this.documentTitle = createElement('label', {
      id: 'documenteditor_title_name',
      styles:
        'font-weight:400;text-overflow:ellipsis;white-space:pre;overflow:hidden;user-select:none;cursor:text',
    });
    let iconCss: string = 'e-de-padding-right';
    let btnFloatStyle: string = 'float:right;';
    let titleCss: string = '';
    if (this.isRtl) {
      iconCss = 'e-de-padding-right-rtl';
      btnFloatStyle = 'float:left;';
      titleCss = 'float:right;';
    }
    // tslint:disable-next-line:max-line-length
    this.documentTitleContentEditor = createElement('div', {
      id: 'documenteditor_title_contentEditor',
      className: 'single-line',
      styles: titleCss,
    });
    this.documentTitleContentEditor.appendChild(this.documentTitle);
    this.tileBarDiv.appendChild(this.documentTitleContentEditor);
    this.documentTitleContentEditor.setAttribute('title', documentTileText);
    let btnStyles: string =
      btnFloatStyle +
      'background: transparent;box-shadow:none; font-family: inherit;border-color: transparent;' +
      'border-radius: 2px;color:inherit;font-size:12px;text-transform:capitalize;height:28px;font-weight:400;margin-top: 2px;';
    // tslint:disable-next-line:max-line-length
    this.close = this.addButton(
      'e-icons e-close e-de-padding-right',
      '',
      btnStyles,
      'de-close',
      closeToolTip,
      false
    ) as Button;
    this.print = this.addButton(
      'e-de-icon-Print ' + iconCss,
      printText,
      btnStyles,
      'de-print',
      printToolTip,
      false
    ) as Button;
    this.open = this.addButton(
      'e-de-icon-Open ' + iconCss,
      openText,
      btnStyles,
      'de-open',
      openText,
      false
    ) as Button;
    let items: ItemModel[] = [
      { text: 'Microsoft Word (.docx)', id: 'word' },
      { text: 'Syncfusion Document Text (.sfdt)', id: 'sfdt' },
    ];
    // tslint:disable-next-line:max-line-length
    this.export = this.addButton(
      'e-de-icon-Download ' + iconCss,
      downloadText,
      btnStyles,
      'documenteditor-share',
      downloadToolTip,
      true,
      items
    ) as DropDownButton;
    if (!isShareNeeded) {
      this.export.element.style.display = 'none';
    } else {
      this.open.element.style.display = 'none';
    }
    if (this.dialogComponent == null) this.close.element.style.display = 'none';
  };
  private setTooltipForPopup(): void {
    // tslint:disable-next-line:max-line-length
    (document
      .getElementById('documenteditor-share-popup') as HTMLElement)
      .querySelectorAll('li')[0]
      .setAttribute(
        'title',
        'Download a copy of this document to your computer as a DOCX file.'
      );
    // tslint:disable-next-line:max-line-length
    (document
      .getElementById('documenteditor-share-popup')as HTMLElement)
      .querySelectorAll('li')[1]
      .setAttribute(
        'title',
        'Download a copy of this document to your computer as an SFDT file.'
      );
  }

  public addUser(actionInfos: ActionInfo | ActionInfo[]): void {
    if (!(actionInfos instanceof Array)) {
        actionInfos = [actionInfos]
    }
    for (let i: number = 0; i < actionInfos.length; i++) {
        let actionInfo: ActionInfo = actionInfos[i];
        if (this.userMap[actionInfo.connectionId as string]) {
            continue;
        }
        let avatar: HTMLElement = createElement('div', { className: 'e-avatar e-avatar-xsmall e-avatar-circle', styles: 'margin: 0px 5px', innerHTML: this.constructInitial(actionInfo.currentUser as string) });
        if (this.userList) {
            this.userList.appendChild(avatar);
        }
        this.userMap[actionInfo.connectionId as string] = avatar;
    }
}

public removeUser(conectionId: string): void {
    if (this.userMap[conectionId]) {
        if (this.userList) {
            this.userList.removeChild(this.userMap[conectionId]);
        }
        delete this.userMap[conectionId];
    }
}

private constructInitial(authorName: string): string {
    const splittedName: string[] = authorName.split(' ');
    let initials: string = '';
    for (let i: number = 0; i < splittedName.length; i++) {
        if (splittedName[i].length > 0 && splittedName[i] !== '') {
            initials += splittedName[i][0];
        }
    }
    return initials;
}
  private wireEvents = (): void => {
    (this.print as Button).element.addEventListener('click', this.onPrint);
    (this.close as Button).element.addEventListener('click', this.onClose);
    (this.open as Button).element.addEventListener('click', (e: Event) => {
      if ((e.target as HTMLInputElement).id === 'de-open') {
        let fileUpload: HTMLInputElement = document.getElementById(
          'uploadfileButton'
        ) as HTMLInputElement;
        fileUpload.value = '';
        fileUpload.click();
      }
    });
    (this.documentTitleContentEditor as HTMLElement).addEventListener(
      'keydown',
      (e): void => {
        if (e.keyCode === 13) {
          e.preventDefault();
          (this.documentTitleContentEditor as HTMLElement).contentEditable = 'false';
          if ((this.documentTitleContentEditor as HTMLElement).textContent === '') {
            (this.documentTitleContentEditor as HTMLElement).textContent = 'Document1';
          }
        }
      }
    );
    (this.documentTitleContentEditor as HTMLElement).addEventListener('blur', (): void => {
      if ((this.documentTitleContentEditor as HTMLElement).textContent === '') {
        (this.documentTitleContentEditor as HTMLElement).textContent = 'Document1';
      }
      (this.documentTitleContentEditor as HTMLElement).contentEditable = 'false';
      (this.documentEditor as DocumentEditor).documentName = (this.documentTitle as HTMLElement).textContent as string;
    });
    (this.documentTitleContentEditor as HTMLElement).addEventListener('click', (): void => {
      this.updateDocumentEditorTitle();
    });
  };
  private updateDocumentEditorTitle = (): void => {
    (this.documentTitleContentEditor as HTMLElement).contentEditable = 'true';
    (this.documentTitleContentEditor as HTMLElement).focus();
    (window.getSelection() as Selection).selectAllChildren((this.documentTitleContentEditor as HTMLElement));
  };
  // Updates document title.
  public updateDocumentTitle = (): void => {
    if (this.documentEditor.documentName === '') {
      this.documentEditor.documentName = 'Untitled';
    }
    (this.documentTitle as HTMLElement).textContent = this.documentEditor.documentName;
  };
  // tslint:disable-next-line:max-line-length
  private addButton(
    iconClass: string,
    btnText: string,
    styles: string,
    id: string,
    tooltipText: string,
    isDropDown: boolean,
    items?: ItemModel[]
  ): Button | DropDownButton {
    let button: HTMLButtonElement = createElement('button', {
      id: id,
      styles: styles,
    }) as HTMLButtonElement;
    this.tileBarDiv.appendChild(button);
    button.setAttribute('title', tooltipText);
    if (isDropDown) {
      // tslint:disable-next-line:max-line-length
      let dropButton: DropDownButton = new DropDownButton(
        {
          select: this.onExportClick,
          items: items,
          iconCss: iconClass,
          cssClass: 'e-caret-hide',
          content: btnText,
          open: (): void => {
            this.setTooltipForPopup();
          },
        },
        button
      );
      return dropButton;
    } else {
      let ejButton: Button = new Button(
        { iconCss: iconClass, content: btnText },
        button
      );
      return ejButton;
    }
  }
  private onPrint = (): void => {
    this.documentEditor.print();
  };
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
  private onExportClick = (args: MenuEventArgs): void => {
    let value: string = args.item.id as string;
    switch (value) {
      case 'word':
        this.save('Docx');
        break;
      case 'sfdt':
        this.save('Sfdt');
        break;
    }
  };
  private save = (format: string): void => {
    // tslint:disable-next-line:max-line-length
    this.documentEditor.save(
      this.documentEditor.documentName === ''
        ? 'sample'
        : this.documentEditor.documentName,
      format as FormatType
    );
  };
}
