import * as React from 'react';
import styles from './ChecklistDocuments.module.scss';
import { IChecklistDocumentsProps } from './IChecklistDocumentsProps';
import { escape, times } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import { DefaultButton, Icon, PrimaryButton } from '@fluentui/react';
import { IItem, Item } from '@pnp/sp/items';
import { Attachment, Attachments, IAttachmentInfo } from '@pnp/sp/attachments';
import { Dialog, IIconProps } from 'office-ui-fabric-react';


export interface IChecklistDocumentsState {
  CheckListRequestData: any;
  requestID: any;
  Notes: any;
  RemoveUploadedFile: any;
  RemoveAttachment: any;
  UploadDocuments: any;
  IsAdmin: boolean;
  IsApproval: boolean;
  IsReviewer: boolean;
  CurrentUserEmail: any;
  ChecklistData: any;
  Isloader: boolean;
  DocumentUploadedSuccessfully : boolean;
}

const navigateBackIcon: IIconProps = { iconName: 'NavigateBack' };

const DocumentUploadedSuccessfullyDialogContentProps = {
};

const documentmodelProps = {
  className : "Successfull-Document"
};


require("../assets/css/style.css");
require("../assets/css/fabric.min.css");

export default class ChecklistDocuments extends React.Component<IChecklistDocumentsProps, IChecklistDocumentsState> {

  constructor(props: IChecklistDocumentsProps, state: IChecklistDocumentsState) {
    super(props);
    this.state = {
      CheckListRequestData: "",
      requestID: "",
      Notes: "",
      RemoveUploadedFile: [],
      RemoveAttachment: [],
      UploadDocuments: [],
      IsAdmin: true,
      IsApproval: false,
      IsReviewer: false,
      CurrentUserEmail: "",
      ChecklistData: "",
      Isloader: false,
      DocumentUploadedSuccessfully : true
    };
  }

  public render(): React.ReactElement<IChecklistDocumentsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <div id="checklistDocuments">
        <div className="ms-Grid">

          <div className='Back'>
            {/* <a className='back-link' href={this.props.context.pageContext._web.absoluteUrl + '/SitePages/Checklist-Request.aspx'} target="_self" data-interception="off"> */}
             <PrimaryButton className="Back-Button" iconProps={navigateBackIcon}  text='Back' href={this.props.context.pageContext._web.absoluteUrl + '/SitePages/Checklist-Request.aspx'} target="_self" data-interception="off"/>
            {/* </a> */}
          </div>

          <div className='ChecklistContainer'>
            <div className='ms-Grid-row ChecklistRequestDetails'>

              <div className='ms-Grid-col ms-sm4 ms-md4 ms-lg4 Checklist'>
                <Icon iconName='CityNext' className="card-Icon mb-10"></Icon>
                <span>Company Name :</span> {this.state.ChecklistData ? this.state.ChecklistData[0].CompanyName : ""}
              </div>

              <div className='ms-Grid-col ms-sm4 ms-md4 ms-lg4 Checklist'>
                <Icon iconName='CheckList' className="card-Icon mb-10"></Icon>
                <span>Checklist Name :</span> {this.state.ChecklistData ? this.state.ChecklistData[0].ChecklistName.ChecklistName : ""}
              </div>

              <div className='ms-Grid-col ms-sm4 ms-md4 ms-lg4 Checklist'>
                <Icon iconName='Mail' className="card-Icon mb-10"></Icon>
                <span>Company Email :</span> {this.state.ChecklistData ? this.state.ChecklistData[0].CompanyEmail : ""}
              </div>

              {/* <div className='ms-Grid-col ms-sm4 ms-md4 ms-lg4 Checklist'>
                <Icon iconName='TextDocument' className="card-Icon mb-10"></Icon>
                <span>Required Documents :</span> {this.state.ChecklistData ? this.state.ChecklistData[0].RequiredDocuments.map(item => item).join(",") : ""}
              </div> */}

            </div>
          </div>

          <div className='Header-Title'>
            <h2 className='Title'>Checklist Documents</h2>
          </div>

          <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '20px' }}>
            <thead>
              <tr style={{ backgroundColor: '#fbfbfb', fontWeight: 'bold' }}>
                <th style={{ padding: '10px', width: '200px', border: '1px solid #ccc' }}>Documents</th>
                <th style={{ padding: '10px', width: '950px', border: '1px solid #ccc' }}>Upload Documents</th>
                <th style={{ padding: '10px', width: '250px', border: '1px solid #ccc' }}>Notes</th>
              </tr>
            </thead>
            <tbody>
              {this.state.CheckListRequestData.length > 0 &&
                this.state.CheckListRequestData.map((item) => (
                  <tr key={item.Id}>
                    <td style={{ padding: '10px', border: '1px solid #ccc' }}>
                      <div className="requiredDocName">{item.Title}</div>
                    </td>

                    <td style={{ padding: '10px', border: '1px solid #ccc' }}>
                      {!this.state.IsAdmin && (
                        <>

                          <label className="Attachmentlabel" htmlFor={item.Title}>
                            Upload Document
                          </label>
                          <input
                            style={{ display: 'none' }}
                            id={item.Title}
                            type="file"
                            multiple
                            onChange={(e) => this.UploadAttachments(e.target.files, item.Id)}
                          />
                        </>
                      )}

                      <div className="Attachment-wrap">
                        {item.Attachments &&
                          item.Attachments.map((i) => (
                            <div className="attachmentname" key={i.FileName}>
                              <a className="link-file" href={'https://200oksolutions.sharepoint.com' + i.ServerRelativeUrl} target="_self" data-interception="off"><p className="file-name">{i.FileName}</p></a>

                              <Icon
                                iconName="Cancel"
                                className="icon-cancel"
                                onClick={() => this.RemoveAttachments(item.Id, i.FileName)}
                              />
                            </div>
                          ))}
                        {item.file &&
                          item.file.map((i) => (
                            <div className="attachmentname" key={i.name}>
                              <p className="file-name">{i.name}</p>
                              <Icon
                                iconName="Cancel"
                                className="icon-cancel"
                                onClick={() => this.RemoveUploadedDoc(item.Id, i.name)}
                              />
                            </div>
                          ))}
                      </div>
                    </td>


                    <td style={{ padding: '10px', border: '1px solid #ccc' }}>{item.Notes}</td>
                  </tr>
                ))}
            </tbody>
          </table>

        </div>

        {
          this.state.Isloader == true ?
            <>
              <div className='LoaderBg-overlay'>
                <div id="loading-wrapper">
                  <div id="loading-text"></div>
                  <div id="loading-content"></div>
                  <label className='Loader-Text'>Please Wait.!!</label>
                </div>
              </div>
            </> : <></>
        }

        <Dialog
          hidden={this.state.DocumentUploadedSuccessfully}
          onDismiss={() =>
            this.setState({
              DocumentUploadedSuccessfully : true
            })
          }
          dialogContentProps={DocumentUploadedSuccessfullyDialogContentProps}
          modalProps={documentmodelProps}
          minWidth={400}
        >
          <div className="confirmation-dialog">
            <div className="checkmark-circle">
            <Icon iconName='CheckMark' className='material-icons'/>
            </div>
            <h2>Awesome!</h2>
            <p>Your Document has been Uploaded Successfully.!!</p>
            <PrimaryButton className="ok-button" onClick={() => this.setState({ DocumentUploadedSuccessfully: true })}>
              OK
            </PrimaryButton>
          </div>
           
        </Dialog>

        {
          !this.state.IsAdmin && (
            <>
              <div className='ms-Grid-row'>
                <div className='Add-Doc'>
                  <div className='ms-Grid-col save-doc'>

                    <PrimaryButton
                      text='Submit'
                      onClick={() => this.AddRequestDocument()}
                    />

                    <div className='Cancel-Doc'>
                      <DefaultButton
                        text='Cancel' onClick={() => this.setState({  })}
                      />
                    </div>
                  </div>
                </div>
              </div>
            </>
          )}

        {
          this.state.IsAdmin && (
            <>
              <div className='ms-Grid-row'>
                <div className='Approve-Doc'>
                  <div className='ms-Grid-col app-doc'>
                    <PrimaryButton
                      text='Approved'
                    />

                    <div className='Reject-Doc'>
                      <PrimaryButton
                        text='Return'
                      />
                    </div>

                    <div className='Return-Doc'>
                      <PrimaryButton
                        text='Rejected'
                      />
                    </div>
                  </div>
                </div>
              </div>
            </>
          )}

      </div>
    );
  }

  public async componentDidMount() {
    await this.GetChecklistDocuments();
    await this.GetCurrentUser();
  }

  public async GetCurrentUser() {
    try {
      const currentUser = await sp.web.currentUser.get();
      const userEmail = currentUser.Email.toLowerCase().trim();
      const ownerGroup = await sp.web.associatedOwnerGroup();
      const groupUsers = await sp.web.siteGroups.getById(ownerGroup.Id).users();

      const isAdmin = groupUsers.some(user =>
        user.LoginName.toLowerCase() === currentUser.LoginName.toLowerCase()
      );

      let groups = await sp.web.currentUser.groups();
      console.log(groups);

      groups.forEach((items) => {
        if (items.Title == "Vivek Owners") {
          this.setState({ IsReviewer: true });
        }
      });
      console.log(this.state.IsReviewer);

      groups.forEach((items) => {
        if (items.Title == "Vivek Owners") {
          this.setState({ IsApproval: true });
        }
      });
      console.log(this.state.IsApproval);

      this.setState({ IsAdmin: isAdmin, CurrentUserEmail: userEmail });
      console.log("IsAdmin:", isAdmin);
    } catch (error) {
      console.error("Error checking admin status:", error);
      this.setState({ IsAdmin: false });
    }
  }

  public async ReivewerControls(Status) {

  }

  public async GetChecklistDocuments() {
    try {
      console.log("Current URL:", window.location.href);
      const urlParams = new URLSearchParams(window.location.search);
      const requestid = urlParams.get("RequestID");
      if (requestid) {
        console.log("RequestID:", requestid);
        this.setState({ requestID: requestid });
        this.GetChecklistRequestDetails(requestid)
        this.GetRequestDocuments(requestid);
      } else {
        console.log("RequestID not found in URL");
      }
    } catch (error) {
      console.error("Error parsing URL parameters:", error);
    }
  }

  public async GetRequestDocuments(ID) {
    const Reqdoc = await sp.web.lists.getByTitle("Request Document").items.select(
      "ID",
      "Title",
      "RequestID/Id",
      "Attachments",
      "Notes",
    ).expand("RequestID").filter(`RequestID/Id eq ${ID}`).orderBy("ID", false).get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(Reqdoc);
      if (data.length > 0) {
        data.forEach(async (item, i) => {

          const item1: IItem = sp.web.lists.getByTitle("Request Document").items.getById(item.Id);
          const info: IAttachmentInfo[] = await item1.attachmentFiles();

          AllData.push({
            Id: item.Id ? item.Id : "",
            Title: item.Title ? item.Title : "",
            RequestID: item.RequestID ? item.RequestID.Id : "",
            Attachments: info,
            isfilechanged: false,
            Notes: item.Notes
          });
          this.setState({ CheckListRequestData: AllData });
          console.log(this.state.CheckListRequestData);
        });
      }
    });
  }

  public async GetChecklistRequestDetails(ID) {

    try {
      const checklist = await sp.web.lists.getByTitle("Checklist Requests").items.select(
        "ID",
        "CompanyName",
        "ChecklistName/ChecklistName",
        "ChecklistName/ID",
        "CompanyEmail",
        "RequiredDocuments",
        "Status"
      ).expand("ChecklistName").filter(`Id eq ${ID}`).get();
      this.setState({ ChecklistData: checklist });
      console.log(this.state.ChecklistData)
    } catch (error) {
      console.log(error);
    }

  }

  // public async GetChecklistRequestlist(ID) {
  //   const checklists = await sp.web.lists.getByTitle("Checklist Requests").items.select(
  //     "ID",
  //     "CompanyName",
  //     "ChecklistName/ChecklistName",
  //     "ChecklistName/ID",
  //     "CompanyEmail",
  //     "RequiredDocuments"
  //   ).expand("ChecklistName").filter(`ID eq ` + ID).get().then((data) => {
  //     let ReqDoc = [];
  //     console.log(data);
  //     console.log(checklists);

  //     let documentDetails = [];

  //     if(data.length > 0 ) {
  //       data.forEach((item) => {
  //         item.RequiredDocuments.forEach((doc) => {
  //           ReqDoc.push({
  //             doc 
  //          });
  //         });

  //       });
  //       this.setState({ CheckListDocumentData : ReqDoc });
  //       console.log(this.state.CheckListDocumentData);
  //     }

  //     if(data.length > 0){
  //       data.forEach((item) => {
  //         documentDetails.push({
  //           ID : item.Id ? item.Id : "",
  //           CompanyName : item.CompanyName ? item.CompanyName : "",
  //           ChecklistName : item.ChecklistName ? item.ChecklistName.ChecklistName : "",
  //           ChecklistNameId : item.ChecklistName ? item.ChecklistName.ID : "",
  //           CompanyEmail : item.CompanyEmail ? item.CompanyEmail : "",
  //           RequiredDocuments : item.RequiredDocuments ? item.RequiredDocuments : ""
  //         });
  //       });
  //       this.setState({ RequestDocumentDetails : documentDetails });
  //       console.log(this.state.RequestDocumentDetails);
  //     }
  //   }).catch((error) => {
  //     console.log(error);
  //   });
  // }

  public async AddRequestDocument() {
    // const addreqdoc = await sp.web.lists.getByTitle("Request Document").items.add({
    //   Notes : this.state.Notes
    // }).catch((error) => {
    //   console.log(error);
    // });

    this.setState({ Isloader: true });

    for (let i = 0; i < this.state.RemoveAttachment.length; i++) {
      const file = this.state.RemoveAttachment[i];

      try {
        const item1: IItem = await sp.web.lists.getByTitle("Request Document").items.getById(file.Id);
        await item1.attachmentFiles.getByName(file.FileName).delete();
        console.log(`Delete file: ${file.FileName}`);
      } catch (eror) {
        console.log(`Error: ${file.FileName}`);
      }
    }
    this.setState({ RemoveAttachment: [] });

    for (let i = 0; i < this.state.UploadDocuments.length; i++) {
      const file = this.state.UploadDocuments[i];

      try {
        const item: IItem = await sp.web.lists.getByTitle("Request Document").items.getById(file.Id);
        await item.attachmentFiles.add(file.FileName, file.file);

        console.log(`Uploaded: ${file.FileName}`);
      } catch (error) {
        console.log(`Error uploading file ${file.FileName}:`, error);
      }

    }
    this.setState({ Isloader: false });
    this.setState({ UploadDocuments: [] });
    this.setState({ DocumentUploadedSuccessfully : false });
  }

  public async UploadAttachments(files, id: number) {
    const updatedChecklist = this.state.CheckListRequestData.map(item => {
      if (item.Id === id) {
        return {
          ...item,
          file: item.file ? [...item.file, ...files] : [...files],
          isfilechanged: true,
        };
      }
      else {
        return item;
      }
    });

    const uploadeddoc = this.state.UploadDocuments;

    const fileArray = [...files];
    fileArray.map(item => {
      uploadeddoc.push({
        Id: id,
        FileName: item.name,
        file: item
      });
    });

    this.setState({ UploadDocuments: uploadeddoc });
    console.log(this.state.UploadDocuments);
    this.setState({ CheckListRequestData: updatedChecklist, UploadDocuments: uploadeddoc });
    console.log("Updated CheckListRequestData:", this.state.CheckListRequestData);
  }

  public async RemoveUploadedDoc(id: number, file) {
    const fileToRemove = file;

    const updatedChecklist = this.state.CheckListRequestData.map(item => {
      if (item.Id === id) {
        const files = Array.isArray(item.file) ? item.file : [item.file]; // force it into array
        const updatedoc = files.filter(f => f.name !== fileToRemove);
        return {
          ...item,
          file: updatedoc,
        };
      }
      return item;
    });

    const updatedoc = this.state.UploadDocuments;
    updatedoc.filter(f => f.FileName !== fileToRemove);

    this.setState({ CheckListRequestData: updatedChecklist, UploadDocuments: updatedoc });
    console.log(this.state.CheckListRequestData);
    console.log(this.state.RemoveUploadedFile);
    // const updatedChecklist = this.state.CheckListRequestData.map(item => {
    //   if (item.Id === id) {
    //   const updatedoc = item.file.filter(f => f.name !== fileToRemove.name);
    //     return {
    //       ...item,
    //       file: updatedoc,
    //     };
    //   }

    //   return item;
    // });
  }

  public async RemoveAttachments(id: number, fileName: string) {
    const updated = this.state.CheckListRequestData.map(item => {
      if (item.Id === id) {
        const updatedFiles = item.Attachments.filter(f => f.FileName !== fileName);
        return {
          ...item,
          Attachments: updatedFiles
        };
      }
      else {
        return item;
      }
    });

    let DeleteDOCS = this.state.RemoveAttachment;
    DeleteDOCS.push({
      FileName: fileName,
      Id: id
    });

    this.setState({ CheckListRequestData: updated, RemoveAttachment: DeleteDOCS });
    console.log(this.state.CheckListRequestData);
    console.log(this.state.RemoveAttachment);
  }

  public async HideNavigation() {

    try {
      // Get current user's groups
      const userGroups = await sp.web.currentUser.groups();

      // Check if the user is in the Owners or Admins group
      const isAdmin = userGroups.some(group =>
        group.Title.indexOf("Owners") !== -1
        ||
        group.Title.indexOf("Admins") !== -1
      );

      if (!isAdmin) {
        // Hide the navigation bar for non-admins
        const navBar = document.querySelector("#SuiteNavWrapper");
        if (navBar) {
          navBar.setAttribute("style", "display: none;");
        }
      } else {
        // Show the navigation bar for admins
        const navBar = document.querySelector("#SuiteNavWrapper");
        if (navBar) {
          navBar.setAttribute("style", "display: block;");
        }
      }
    } catch (error) {
      console.error("Error checking user permissions: ", error);
    }

  }

}
