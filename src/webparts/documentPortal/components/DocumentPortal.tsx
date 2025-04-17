import * as React from 'react';
import styles from './DocumentPortal.module.scss';
import { IDocumentPortalProps } from './IDocumentPortalProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton, DetailsList, Dialog, Dropdown, IColumn, Icon, IIconProps, Label, PrimaryButton, SearchBox, TextField } from 'office-ui-fabric-react';
import { sp } from '@pnp/sp';
import { Item } from '@pnp/sp/items';



export interface IDocumentPortalState {
  CheckListRequestData : any;
  AllChecklistRequestData : any;
  CompanyName : any;
  ChecklistNameID : any;
  ChecklistName : any;
  CompanyEmail : any;
  RequiredDocuments : any;
  AddRequestDialog: boolean;
  ChecklistRequestlist : any;
  RequiredDocumentslist : any;
  searchText :  any;
  RequestsData : any;
  FilteredData : any;
  ChecklistDocuments : any;
  AddCustomDocDialog : boolean;
  CustomDocumentFiled : any;
  CustomDocumentsSave : any;
  Title : any;
  SelectedChecklistDocument : any;
  EditCompanyName : any;
  EditChecklistName : any;
  EditCompanyEmail : any;
  EditRequiredDocuments : any;
  EditChecklistNameID : any;
  EditSelectedChecklistDocument : any;
  EditChecklistRequestDialog: boolean;
  CurrentChecklistRequestID : any;
  DeleteSelectedChecklistID : any;
  DeleteChecklistRequestDialog : boolean;
  ProjectChecklistID : any;
  EditAddedTag : any;
  IsAdmin : boolean;
  CurrentUserEmail : any;
  MyChecklistRequestData : any;
}

const addIcon: IIconProps = { iconName: 'Add' };

const SendIcon : IIconProps = { iconName: 'Send'};

const CancelIcon : IIconProps = { iconName: 'Cancel'};

const DeleteIcon : IIconProps = { iconName: 'Delete'};

const TextDocumentEdit : IIconProps = { iconName: 'TextDocumentEdit' };

const AddChecklistRequestDialogContentProps = {
  title: "Add Checklist Details",
};

const AddChecklistDocuDialogmentContentProps = {
  title: "",
};

const ReadChecklistRequestDialogContentProps = {
  title: "Read Checklist Details"
};

const UpdateChecklistRequestDialogContentProps = {
  title: "Update Checklist Details"
};

const DeleteChecklistRequestFilterDialogContentProps = {
  title:"Confirm Deletion"
};

const adddocumntsaveProps = {
  className: "Add-Save"
};

const addmodelProps = {
  className: "Add-Dialog"
};

const readmodelProps = {
  className: "Read-Dialog"
};

const updatemodelProps = {
  className : "Update-Dialog"
};

const deletmodelProps = {
  className : "Delet-Dialog"
};

require("../assets/css/style.css");
require("../assets/css/fabric.min.css");


export default class DocumentPortal extends React.Component<IDocumentPortalProps, IDocumentPortalState> {

  constructor(props: IDocumentPortalProps, state: IDocumentPortalState) {

    super(props);

    this.state = {
      CheckListRequestData : "",
      AllChecklistRequestData : [],
      CompanyName : "",
      ChecklistNameID : "",
      ChecklistName : "",
      CompanyEmail : "",
      RequiredDocuments : "",
      AddRequestDialog : true,
      ChecklistRequestlist : [],
      RequiredDocumentslist : [],
      searchText : [],
      RequestsData : "",
      FilteredData : "",
      ChecklistDocuments : [],
      AddCustomDocDialog : true,
      CustomDocumentFiled : "",
      CustomDocumentsSave : [],
      Title : "",
      SelectedChecklistDocument : [],
      EditCompanyName : "",
      EditChecklistName : "",
      EditCompanyEmail : "",
      EditChecklistNameID : [],
      EditRequiredDocuments : "",
      EditSelectedChecklistDocument : [],
      EditChecklistRequestDialog :true, 
      CurrentChecklistRequestID : "",
      DeleteSelectedChecklistID : "",
      DeleteChecklistRequestDialog : true,
      ProjectChecklistID : "",
      EditAddedTag : "",
      IsAdmin : true,
      CurrentUserEmail : "",
      MyChecklistRequestData : []
    };

  }

  public render(): React.ReactElement<IDocumentPortalProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;


   const columns : IColumn[] = [
      {
        key: "CompanyName",
        name: "Company Name",
        fieldName: "CompanyName",
        minWidth: 170,
        maxWidth : 170,
        isResizable : false
      },
      {
        key: "ChecklistName",
        name: "Checklist Name",
        fieldName: "ChecklistName",
        minWidth: 150,
        maxWidth : 150,
        isResizable : false,
        // onRender: (item) => (<div>{item["ChecklistName"]["Title"]}</div>) 
      },
      {
        key: "CompanyEmail",
        name: "Company Email",
        fieldName: "CompanyEmail",
        minWidth: 220,
        maxWidth : 220,
        isResizable : false
      },
      {
        key: "Status",
        name: "Status",
        fieldName: "Status",
        minWidth: 130,
        maxWidth : 130,
        isResizable : false
      },
      {
        key: "RequiredDocuments",
        name: "Required Documents",
        fieldName: "RequiredDocuments",
        minWidth: 270,
        maxWidth :270,
        isResizable : false,
        onRender : (item) =>  {
          return (
            <div>
              {item.RequiredDocuments.map((value) => (
                <div className='Document-Tag'>
                  {value}
                </div>
              ))}
            </div>
          );
        }
        
      },
      {
        key : "Actions",
        name: "Actions",
        fieldName : "",
        minWidth : 150,
        maxWidth : 150,
        isResizable : false,
        onRender : (item) => {
          return(
            <div>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col">
                  <div className='Request-Icon'>

                    <div className="Read-Icon">
                      <a href={this.props.context.pageContext._web.absoluteUrl + '/SitePages/Checklist-Documents.aspx?RequestID=' + item.ID} target="_self" data-interception="off">
                        <Icon iconName="View" className='read-doc'></Icon>
                      </a>
                      
                    </div>
                    {
                      this.state.IsAdmin && (
                        <>
                          <div className='Edit-Icon'>
                            <Icon className='Edit-Icon' iconName='Edit' onClick={() => this.setState({ EditChecklistRequestDialog : false , CurrentChecklistRequestID : item.ID , ProjectChecklistID : item.ID}, () => this.GetEditChecklistRequest(item.ID))}></Icon>
                          </div>

                          <div className="Delete-Icon">
                              <Icon className='icon' iconName="Delete" onClick={() => this.setState({ DeleteChecklistRequestDialog : false , DeleteSelectedChecklistID : item.ID})}></Icon>
                          </div>
                        </>
                      )}

                  </div>
                </div>
              </div>
            </div>  
          );
          
        }
      }
    ];

    return (
      
        <div id="documentPortal">
          <div className='ms-Grid'>

          <div className='Header-Title'>
            <h2 className='Title'>Checklist Request Management</h2>
          </div>
            
            <div className="ms-Grid-row">
              <div className="filedGroup">

                <div className="ms-Grid-col ms-sm5 ms-md4 ms-lg2">  
                    <SearchBox placeholder="Search" className="new-search" 
                      onChange={(e) => {this.applyVendorFilters(e.target.value);}}
                      onClear={(e) => {this.applyVendorFilters(e.target.value);}} 
                    />
                </div>
              
                { 
                  this.state.IsAdmin &&
                    <div className='ms-Grid-col ms-sm1 ms-md1 ms-lg10 Add-Checklist'>
                      <div className='Add-Request'>
                        <PrimaryButton iconProps={addIcon} text="Add Request" onClick={() => this.setState({ AddRequestDialog: false })}/>
                      </div>
                    </div>
                }
              </div>
            </div>
          </div>

          <Dialog
            hidden={this.state.AddRequestDialog}
            onDismiss={() =>
              this.setState({
                
                CompanyName : "",
                ChecklistName : "",
                ChecklistNameID : "",
                CompanyEmail : "",
                RequiredDocuments : "",
                SelectedChecklistDocument : "",
                AddRequestDialog : true,
              })
            }
              dialogContentProps={AddChecklistRequestDialogContentProps}
              modalProps={addmodelProps}
              minWidth={600}
            >

            <div className="ms-Grid-row">
              
              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6 mb-5'>
                <div className="Add-CompanyName">
                  <TextField
                    label="Company Name"
                    type="Text"
                    required={true}
                    onChange={(value) =>
                      this.setState({ CompanyName : value.target["value"]})
                    }
                  />
                </div>
              </div>

              <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
                <div className="Add-CompanyEmail">
                    <TextField
                      label="Enter your company Email-Address"
                      placeholder='your@email.com'
                      type='Email'
                      required={true}
                      onChange={(value) =>
                        this.setState({ CompanyEmail : value.target["value"]})
                      }
                     
                    />
                </div>
              </div>

              <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
                <div className="Add-Checklistname">
                  <Dropdown
                    options={this.state.ChecklistRequestlist}
                    label="Checklist Name"
                    required
                    placeholder="Select your Checklist Name"
                    onChange={(e, option, text) =>
                      this.setState({ ChecklistName : option.text, ChecklistNameID : option.key }, () => this.handleChecklistName(option.text))
                    }    
                  />
                </div>
              </div>

              <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg12">

                    <div className='Checklist-Documentheader'>
                      <Label className='lable-title'>Checklist Documents :</Label>
                        <div className='Add-SaveDoc'>
                            <PrimaryButton 
                              type='Add' text='Add-CustomDocument' onClick={() => this.setState({ AddCustomDocDialog : false })} 
                            />
                        </div>
                    </div>
                      
                      <div className="Add-RequiredDocument">
                        {
                          this.state.SelectedChecklistDocument.length > 0 && 
                            this.state.SelectedChecklistDocument.map((item, ID) => {
                              return (
                                <>
                                      <div className='DocTag'>
                                      {item.Title}
                                      <Icon iconName='Cancel' className='icon-cancel' onClick={() => this.UnSaveDocument(item.Title)}/>
                                      </div>
                                </>
                              );
                            }
                          )
                        } 
                  </div>
              </div>

              <Dialog 
                hidden={this.state.AddCustomDocDialog}
                onDismiss={() =>
                  this.setState({
                    CustomDocumentsSave : "",
                    AddCustomDocDialog : true
                  })
                }
                dialogContentProps={AddChecklistDocuDialogmentContentProps}
                modalProps={adddocumntsaveProps}
                minWidth={400}
                >
                <div>
                  <div className='ms-Grid-row'>
                    <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                      <TextField
                        label='Add-CustomDocument'
                        type='Text'
                        required={true}
                        onChange={(value) => 
                          this.setState({ CustomDocumentsSave : value.target["value"]})
                        }
                      />
                    </div>
                  </div>
                </div>

                <div className='ms-Grid-row'>
                  <div className='Add-custom'>
                    <div className='ms-Grid-col save-doocument'>
                          <PrimaryButton
                            text='Save'
                            onClick={() => this.SaveDocuments()}
                          />
                        <div className='Cancel-Doc'>
                          <DefaultButton 
                              text='Cancel'
                              onClick={() => this.setState({ AddCustomDocDialog : true })}
                          />
                        </div>
                        
                    </div>
                  </div>
                </div>
              </Dialog>
            </div>

            <div className='ms-Grid-row'>
              <div className='Add-RequestCheck'>
              
                  <div className='ms-Grid-col Add-Submit'>
                      <PrimaryButton
                        iconProps={SendIcon}
                        text="Submit"
                        onClick={() => this.AddChecklistRequest()}
                      />
                  </div>
                

                <div className='ms-Grid-col Cancel-Request'>
                  <DefaultButton
                    iconProps={CancelIcon}
                    text="Cancel"
                    onClick={() => this.setState({ AddRequestDialog :  true })}
                  />
                </div>
              </div>
            </div>
          </Dialog>

          <Dialog
            hidden={this.state.EditChecklistRequestDialog}
            onDismiss={() =>
              this.setState({
                EditChecklistRequestDialog : true,
                EditCompanyName : "",
                EditCompanyEmail : "",
                EditChecklistName : "",
                EditRequiredDocuments : "",
                EditSelectedChecklistDocument : ""
              })
            }
            dialogContentProps={UpdateChecklistRequestDialogContentProps}
            modalProps={updatemodelProps}
            minWidth={600}
            >
              <div>
                <div className='ms-Grid-row'>

                  <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                    <div className='Edit-CompanyName'>
                      <TextField
                        label='Company Name'
                        type='Text'
                        required={true}
                        onChange={(value) =>
                          this.setState({ EditCompanyName : value.target["value"]})
                        }
                        value={this.state.EditCompanyName}
                      />
                    </div>
                  </div>

                  <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                    <div className='Edit-CompanyEmail'>
                      <TextField
                        label='Company Email'
                        type='Text'
                        required={true}
                        onChange={(value) =>
                          this.setState({ EditCompanyEmail : value.target["value"]})
                        }
                        value={this.state.EditCompanyEmail}
                      />
                    </div>
                  </div>

                  <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
                    <div className="Add-Checklistname">
                      <Dropdown
                        options={this.state.ChecklistRequestlist}
                        label="Checklist Name"
                        required
                        placeholder="Select your Checklist Name"
                        defaultSelectedKey={this.state.EditChecklistNameID}
                        onChange={(e, option, text) =>
                          this.setState({ EditChecklistNameID : option.text  })
                        }    
                      />
                    </div>
                  </div>    

                  <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg12">

                            <div className='Checklist-Documentheader'>
                              <Label className='lable-title'>Checklist Documents :</Label>
                                <div className='Add-SaveDoc'>
                                    <PrimaryButton 
                                      type='Add' text='Add-CustomDocument' onClick={() => this.setState({ AddCustomDocDialog : false })} 
                                    />
                                </div>
                            </div>
                              
                              <div className="Add-RequiredDocument">
                                {
                                  this.state.SelectedChecklistDocument.length > 0 && 
                                    this.state.SelectedChecklistDocument.map((item, ID) => {
                                      return (
                                        <>
                                              <div className='DocTag'>
                                              {item.Title}
                                              <Icon iconName='Cancel' className='icon-cancel' onClick={() => this.UnSaveDocument(item.Title)}/>
                                              </div>
                                        </>
                                      );
                                    }
                                  )
                                } 
                            </div>
                  </div>

                    <Dialog 
                      hidden={this.state.AddCustomDocDialog}
                      onDismiss={() =>
                        this.setState({
                          AddCustomDocDialog : true,
                          CustomDocumentsSave : "",
                  
                        })
                      }
                      dialogContentProps={AddChecklistDocuDialogmentContentProps}
                      modalProps={adddocumntsaveProps}
                      minWidth={400}
                    >
                      <div>
                        <div className='ms-Grid-row'>
                          <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                            <TextField
                              label='Add-CustomDocument'
                              type='Text'
                              required={true}
                              onChange={(value) => 
                                this.setState({ CustomDocumentsSave : value.target["value"]})
                              }
                            />
                          </div>
                        </div>
                      </div>

                      <div className='ms-Grid-row'>
                        <div className='Add-custom'>
                          <div className='ms-Grid-col save-doocument'>
                                <PrimaryButton
                                  text='Save'
                                  onClick={() => this.SaveDocuments()}
                                />
                              <div className='Cancel-Doc'>
                                <DefaultButton 
                                    text='Cancel'
                                    onClick={() => this.setState({ AddCustomDocDialog : true })}
                                />
                              </div>
                              
                          </div>
                        </div>
                      </div>
                    </Dialog>

                </div>
              </div>

              <div className='ms-Grid-row'>
                <div className='Edit-RequestCheck'>
                  <div className='ms-Grid-col Edit-Submit'>
                      <PrimaryButton
                        iconProps={TextDocumentEdit}
                        text="Update"
                        onClick={() => this.UpdateChecklistRequest(this.state.CurrentChecklistRequestID)}
                      />
                  </div>

                  <div className='ms-Grid-col Edit-Cancel-Request'>
                    <DefaultButton
                      iconProps={CancelIcon}
                      text="Cancel"
                      onClick={() => this.setState({ EditChecklistRequestDialog : true })}
                    />
                  </div>

                </div>
              </div>

          </Dialog>

          <Dialog
            hidden={this.state.DeleteChecklistRequestDialog}
            onDismiss={() =>
              this.setState({
                DeleteChecklistRequestDialog : true
              })
            }
            dialogContentProps={DeleteChecklistRequestFilterDialogContentProps}
            modalProps={deletmodelProps}
            minWidth={500}
          >

            <div className="DeleteClose-Icon">
              <div className='delete-text'>
                {/* <h5 className='confirm-text'>Confirm Deletion</h5> */}
                <Icon iconName="Cancel" className='confirm-icon' onClick={() => this.setState({ DeleteChecklistRequestDialog : true })}></Icon>
              </div>
              <div className="delete-msg">
                <Icon iconName='Warning' className='Warinig-Ic'></Icon>
                <p className='mb-0'>Are you sure? <br/>Do you really want to delete these record? </p>
              </div>
              <div className='Delet-buttons'>
                <DefaultButton
                  className="cancel-Icon"
                  text='Cancel'
                  iconProps={CancelIcon}
                  onClick={() => this.setState({ DeleteChecklistRequestDialog : true })}
                />

                <PrimaryButton
                  className='delete-icon'
                  text='Delete'
                  iconProps={DeleteIcon}
                  onClick={() => this.DeleteChecklists(this.state.DeleteSelectedChecklistID)}
                />
              </div>
            </div>

          </Dialog>

          <div className='ms-Grid'>
            <DetailsList
              className='checklistrequest-List'
              items={this.state.CheckListRequestData}
              columns={columns}
              setKey="set"
              layoutMode={1}
              selectionMode={0}
              isHeaderVisible={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="select row"
            >
            </DetailsList>
          </div>

        </div>
    );
  }

  public async componentDidMount() {
    this.GetProjectChecklistName();
    this.GetchecklistRequest();
    this.GetChecklistDocuments();
    this.GetChecklistRequestItem();
    this.GetCurrentUser();
    this.HideNavigation();
  }

  // public async GetCurrentUser() {
  //   let ownerGroup = await sp.web.associatedOwnerGroup();
  //   console.log(ownerGroup);

  //     if(ownerGroup.OwnerTitle) {
  //       this.setState({ IsAdmin : false });
  //     }
     
  //    console.log(this.state.IsAdmin);
  // }

  public async GetCurrentUser() {
    try {
      const currentUser = await sp.web.currentUser.get();
      const userEmail = currentUser.Email.toLowerCase().trim();
      const ownerGroup = await sp.web.associatedOwnerGroup();
      const groupUsers = await sp.web.siteGroups.getById(ownerGroup.Id).users();
  
      const isAdmin = groupUsers.some(user =>
        user.LoginName.toLowerCase() === currentUser.LoginName.toLowerCase()
      );
  
      this.setState({ IsAdmin: isAdmin, CurrentUserEmail : userEmail });
      console.log("IsAdmin:", isAdmin);
    } catch (error) {
      console.error("Error checking admin status:", error);
      this.setState({ IsAdmin: false }); 
    }
  }
  

  public  onChange = (event, option) => {
    let selectedItems = this.state.RequiredDocuments;
    if (option.selected) {
      selectedItems.push(option.key); // Add to the selected items array
    } else {
      // Remove from the selected items array
      (selectedItems => selectedItems.filter(item => item !== option.key));
    }
    this.setState({ RequiredDocuments: selectedItems });
  }

  public async GetProjectChecklistName () {
    const requestItem = await sp.web.lists.getByTitle("Project Checklists").items.select(
      "ID",
      "ChecklistName"
    ).get().then((data) => {
      let RequestData = [];
      data.forEach(function (dname ,i) {
        RequestData.push({ key : dname.ID , text: dname.ChecklistName });
      });
      console.log(requestItem);
      this.setState({ ChecklistRequestlist : RequestData });
    })
    .catch((error) => {
      console.log(error);
    });
  }

 public async GetChecklistDocuments() {
    const documentlist = await sp.web.lists.getByTitle("Checklist Documents").items.select(
      "ID",
      "Title",
      "ChecklistName/ChecklistName",
      "ChecklistName/ID"
    ).expand("ChecklistName").get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(documentlist);

      if(data.length > 0){
        data.forEach((item) => {
          AllData.push({
            ID : item.Id ? item.Id : "",
            Title : item.Title ? item.Title : "",
            ChecklistName : item.ChecklistName ? item.ChecklistName.ChecklistName : "",
          });
        });
        this.setState({ ChecklistDocuments : AllData });
        console.log(this.state.ChecklistDocuments);
      }
    }).catch((error) => {
      console.log(error); 
    });
  }

  // public async GetchecklistRequest() {
  //   const request = await sp.web.lists.getByTitle("Checklist Requests").items.select(
  //     "ID",
  //     "CompanyName",
  //     "ChecklistName/ChecklistName",
  //     "ChecklistName/ID",
  //     "CompanyEmail",
  //     "RequiredDocuments",
  //     "Status"
  //   ).expand("ChecklistName").get().then((data) => {
  //     let AllData = [];
  //     console.log(data);
  //     console.log(request);

  //     if(data.length > 0){
  //       data.forEach((item) => {
  //         AllData.push({
  //           ID : item.Id ? item.Id : "",
  //           CompanyName : item.CompanyName ? item.CompanyName : "",
  //           ChecklistName : item.ChecklistName ? item.ChecklistName.ChecklistName : "",
  //           ChecklistNameId : item.ChecklistName ? item.ChecklistName.ID : "",
  //           CompanyEmail : item.CompanyEmail ? item.CompanyEmail : "",
  //           RequiredDocuments : item.RequiredDocuments ? item.RequiredDocuments : "",
  //           Status : item.Status ? item.Status : ""
  //         });
  //       });
  //       this.setState({ CheckListRequestData : AllData });
  //       console.log(this.state.CheckListRequestData);
  //     }
  //   }).catch((error) => {
  //     console.log(error);
  //   });
  // }

  public async GetchecklistRequest() {
    try {
      const currentUser = await sp.web.currentUser.get();
      const currentUserEmail = currentUser.Email.toLowerCase();

      const isAdmin = await this.GetCurrentUser(); 

      let checklistItems = [];
  
      if (this.state.IsAdmin) {
        checklistItems = await sp.web.lists.getByTitle("Checklist Requests").items
          .select(
            "ID",
            "CompanyName",
            "ChecklistName/ChecklistName",
            "ChecklistName/ID",
            "CompanyEmail",
            "RequiredDocuments",
            "Status"
          )
          .expand("ChecklistName")
          .get();
      } else {
        checklistItems = await sp.web.lists.getByTitle("Checklist Requests").items
          .filter(`CompanyEmail eq '${currentUserEmail}'`)
          .select(
            "ID",
            "CompanyName",
            "ChecklistName/ChecklistName",
            "ChecklistName/ID",
            "CompanyEmail",
            "RequiredDocuments",
            "Status"
          )
          .expand("ChecklistName")
          .get();
      }
  
      const AllData = checklistItems.map((item) => ({
        ID : item.Id ? item.Id : "",
        CompanyName : item.CompanyName ? item.CompanyName : "",
        ChecklistName : item.ChecklistName ? item.ChecklistName.ChecklistName : "",
        ChecklistNameId : item.ChecklistName ? item.ChecklistName.ID : "",
        CompanyEmail : item.CompanyEmail ? item.CompanyEmail : "",
        RequiredDocuments : item.RequiredDocuments ? item.RequiredDocuments : "",
        Status : item.Status ? item.Status : ""
      }));
  
      this.setState({ CheckListRequestData: AllData , AllChecklistRequestData : AllData });
      console.log(this.state.CheckListRequestData);
    } catch (error) {
      console.error("Error fetching checklist requests:", error);
    }
  }
  

  public async GetChecklistRequestItem() {
    const choiceFieldName2 = "Required Documents";
    const field2 = await sp.web.lists.getByTitle("Checklist Requests").fields.getByInternalNameOrTitle(choiceFieldName2)();
    let requireddocuments = [];
    field2["Choices"].forEach(function (dname,i) {
      requireddocuments.push({ key : dname , text : dname });
    });
    console.log(field2);
    this.setState({ RequiredDocumentslist : requireddocuments });
  }

  public async handleChecklistName(SelectedChecklistName) {
    let checklist = this.state.ChecklistDocuments;
    const selectedChecklist = checklist.filter((item) => {
      if(item.ChecklistName == SelectedChecklistName) {
        return item;
      }
    }
  );
    console.log(selectedChecklist);
    this.setState({ SelectedChecklistDocument : selectedChecklist });
  }

  public async SaveDocuments() {
    let documents = this.state.SelectedChecklistDocument;
    documents.push({
      ID : "",
      Title : this.state.CustomDocumentsSave,
      ChecklistName : this.state.ChecklistName
    });
    this.setState({ SelectedChecklistDocument : documents });
    this.setState({ AddCustomDocDialog : true });
  }

  public async UnSaveDocument(titleToRemove: string) {
    let documents = this.state.SelectedChecklistDocument;
  
    let updatedDocuments = documents.filter(doc => doc.Title !== titleToRemove);

    this.setState({ 
      SelectedChecklistDocument: updatedDocuments,
    });
  }
  

  public async AddChecklistRequest() {
      if(this.state.CompanyName.length == 0) {
        alert("Please Complete Details.");
      }
      else
      {
          const addRequest : any = await sp.web.lists.getByTitle("Checklist Requests").items.add({
            CompanyName : this.state.CompanyName,
            ChecklistNameId : this.state.ChecklistNameID,
            CompanyEmail : this.state.CompanyEmail,
            RequiredDocuments : { results : this.state.SelectedChecklistDocument.map((doc) => doc.Title)} 
        }).catch((error) => {
          console.log(error);
        });

        let requestDoc = "Request Document";
        
        this.state.SelectedChecklistDocument.forEach((item) => {
          sp.web.lists.getByTitle(requestDoc).items.add({
            Title : item.Title,
            RequestIDId : addRequest.data.ID
          }).then((data) => {
            console.log(data);
          }).catch((error) => {
            console.log(error);
          });
        });


        this.GetchecklistRequest();
        this.setState({ CheckListRequestData : addRequest });
        this.setState({ AddRequestDialog : true });
      }
  }

  private async applyVendorFilters(Test)
  {
    if(Test)
    {
      let SerchText = Test.toLowerCase();

    let filteredData = this.state.AllChecklistRequestData.filter((x) => {
      let CompanyName = x.CompanyName.toLowerCase();
      let CompanyEmail = x.CompanyEmail.toLowerCase();
      return(
        CompanyName.includes(SerchText) || CompanyEmail.includes(SerchText)
      );
    });

    this.setState({ CheckListRequestData:filteredData });
    }
    else
    {
      this.setState({ CheckListRequestData:this.state.AllChecklistRequestData });
    }
  }

  public async GetEditChecklistRequest(ID) {
    let Editchecklistrequest = this.state.CheckListRequestData.filter((item) => {
      if(item.ID == ID) {
        return item;
      }
    });

    let editDocdata = [];
    if(Editchecklistrequest[0].RequiredDocuments.length > 0){
      Editchecklistrequest[0].RequiredDocuments.forEach((item) => {
        editDocdata.push({
          ID : "",
          Title : item,
          ChecklistName : Editchecklistrequest[0].ChecklistName
        });
      });
      this.setState({ SelectedChecklistDocument : editDocdata });
      console.log(this.state.SelectedChecklistDocument);
    }
    console.log(Editchecklistrequest);



    this.setState({
      EditCompanyName : Editchecklistrequest[0].CompanyName,
      EditCompanyEmail : Editchecklistrequest[0].CompanyEmail,
      EditChecklistNameID : Editchecklistrequest[0].ChecklistNameId,
    });
  }

  public async UpdateChecklistRequest(CurrentChecklistRequestID) {

    const updaterquestlist = await sp.web.lists.getByTitle("Checklist Requests").items.getById(CurrentChecklistRequestID).update({
      CompanyName : this.state.EditCompanyName,
      ChecklistNameId : this.state.EditChecklistNameID,
      CompanyEmail : this.state.EditCompanyEmail,
      RequiredDocuments : {results : this.state.SelectedChecklistDocument.map((doc) => doc.Title)}
    }).catch((error) => {
      console.log(error);
    });



    this.setState({ EditChecklistRequestDialog : true });
    this.setState({ CheckListRequestData : updaterquestlist });
    this.GetchecklistRequest();
  } 

  public async DeleteChecklists(DeleteSelectedChecklistID) {
    const deletechecklist = await sp.web.lists.getByTitle("Checklist Requests").items.getById(DeleteSelectedChecklistID).delete();
    this.setState({ CheckListRequestData : deletechecklist});
    this.setState({ DeleteChecklistRequestDialog : true });
    this.GetchecklistRequest();
  }

  public async HideNavigation(){
   
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
