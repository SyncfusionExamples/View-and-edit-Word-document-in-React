import React from 'react';
import './App.css';

import {
  DocumentEditorContainerComponent, Toolbar
} from '@syncfusion/ej2-react-documenteditor';
import { TitleBar } from './title-bar';
import { isNullOrUndefined } from '@syncfusion/ej2-base';

DocumentEditorContainerComponent.Inject(Toolbar);

export class Default extends React.Component<{}, {}> {
  titleBar: TitleBar;
  container: DocumentEditorContainerComponent;

  onCreated(): void {
    let titleBarElement: HTMLElement = document.getElementById('default_title_bar') as HTMLElement;
    this.titleBar = new TitleBar(titleBarElement, this.container.documentEditor, true);
    this.container.documentEditor.documentName = 'Getting Started';
    this.titleBar.updateDocumentTitle();
    //Sets the language id as EN_US (1033) for spellchecker and docker image includes this language dictionary by default.
    //The spellchecker ensures the document content against this language.
    this.container.documentEditor.spellChecker.languageID = 1033;
    this.container.documentChange = function () {
      if (!isNullOrUndefined(this.titleBar)) {
        this.titleBar.updateDocumentTitle();
      }
      this.container.documentEditor.focusIn();
    }
    setInterval(() => {
      this.updateDocumentEditorSize();
    }, 100);
    //Adds event listener for browser window resize event.
    window.addEventListener("resize", this.onWindowResize);

    this.openTemplate();
  }
  onWindowResize = (): void => {
    //Resizes the document editor component to fit full browser window automatically whenever the browser resized.
    this.updateDocumentEditorSize();
  }
  updateDocumentEditorSize(): void {
    //Resizes the document editor component to fit full browser window.
    var windowWidth = window.innerWidth;
    //Reducing the size of title bar, to fit Document editor component in remaining height.
    var windowHeight = window.innerHeight - this.titleBar.getHeight();
    this.container.resize(windowWidth, windowHeight);
  }

  openTemplate(): void {
    var uploadDocument = new FormData();
    uploadDocument.append('DocumentName', 'Getting Started.docx');
    var loadDocumentUrl = this.container.serviceUrl + 'LoadDocument';
    var httpRequest = new XMLHttpRequest();
    httpRequest.open('POST', loadDocumentUrl, true);
    var dataContext = this;
    httpRequest.onreadystatechange = function () {
      if (httpRequest.readyState === 4) {
        if (httpRequest.status === 200 || httpRequest.status === 304) {
          //Opens the SFDT for the specified file received from the web API.
          dataContext.container.documentEditor.open(httpRequest.responseText);
        }
      }
    };
    //Sends the request with template file name to web API. 
    httpRequest.send(uploadDocument);
  }

  render() {
    return (
      <div>
        <div id="default_title_bar" className="e-de-ctn-title"></div>
        <DocumentEditorContainerComponent id="container" ref={(scope) => { this.container = scope; }} height={'590px'} serviceUrl={"http://localhost:6002/api/documenteditor/"} enableToolbar={true} enableSpellCheck={true} created={this.onCreated.bind(this)} />
      </div>
    );
  }
}
