import * as React from "react";
import { Stack, TextField, Dropdown, IDropdownOption, PrimaryButton, IPersonaProps, Label, DefaultButton, } from "@fluentui/react";
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import Swal from "sweetalert2";
import { IAddTaskProps } from "./IAddTaskProps";
import { TaskService } from "../../../Services/TaskService";
import styles from "./AddTask.module.scss";

interface IAddTaskState {
  Title: string;
  Description: string;
  DueDate: string;
  Priority: {
    Id: string;
    Title: string;
  };
  Status: string;
  Category: string;
  AssignedBy: {
    Id?: string;
    Title?: string;
  };
  peoplePickerItems: IPersonaProps[];
  isSubmitting: boolean;
  errorMessage: string;
  fieldErrors: { [key in keyof IAddTaskState]?: string };
}


export default class AddTask extends React.Component<IAddTaskProps, IAddTaskState> {

  constructor(props: IAddTaskProps) {
    super(props);
    this.state = {
      Title: "",
      Description: "",
      DueDate: "",
      Priority: {
        Id: "",
        Title: ""
      },
      Status: "",
      Category: "",
      AssignedBy: {
        Id: "",
        Title: ""
      },
      peoplePickerItems: [],
      isSubmitting: false,
      errorMessage: "",
      fieldErrors: {
        Title: "",
        Description: "",
        DueDate: "",
        Priority: "",
        Status: "",
        Category: "",
        AssignedBy: "",
      },
    };
  }


  handleInputChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    const { name } = event.currentTarget;
    this.setState({
      [name]: newValue || "",
      fieldErrors: { ...this.state.fieldErrors, [name]: "" }, 
    } as unknown as Pick<IAddTaskState, keyof IAddTaskState>);
  };

  handleDropdownChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    const { key, text } = option || {};
    const fieldName = (event.target as HTMLDivElement).getAttribute("data-field-name");
  
    if (fieldName) {
      this.setState((prevState) => {
        let updatedField: string | number | undefined | { Id: string, Title: string };
  
        if (fieldName === "Priority") {
          updatedField = { Id: key as string, Title: text || "" };
        } else if (fieldName === "Status") {
          updatedField = key as string;
        } else {
          updatedField = key;
        }
  
        return {
          ...prevState,
          [fieldName as keyof IAddTaskState]: updatedField,
          fieldErrors: {
            ...prevState.fieldErrors,
            [fieldName]: key ? "" : "Please select a value",
          },
        };
      });
    }
  };
  
  
  handlePeoplePickerChange = (items: IPersonaProps[]): void => {
    if (items.length > 0) {
      this.setState({
        peoplePickerItems: items,
        AssignedBy: {
          Id: items[0].id, 
          Title: items[0].text 
        },
      });
    } else {
      this.setState({
        AssignedBy: {
          Id: "",
          Title: ""
        }
      });
    }
  };

  validateForm = (): boolean => {
    
    const { Title, Description, DueDate, Priority, Status, Category, AssignedBy } = this.state;
    let isValid = true;
    const errors = { ...this.state.fieldErrors };
    const today = new Date();
    today.setHours(0,0,0,0)

    if (!Title.trim()) {
      errors.Title = "Title is required.";
      isValid = false;
    }
    if (!Description.trim()) {
      errors.Description = "Description is required.";
      isValid = false;
    }
    if (!DueDate) {
      errors.DueDate = "Due Date is required.";
      isValid = false;
    } else {
      const dueDate = new Date(DueDate);
      if (dueDate < today) {
        errors.DueDate = "Due Date cannot be in the past.";
        isValid = false;
      }
    }
    if (!Priority.Title) {
      errors.Priority = "Priority is required.";
      isValid = false;
    }
    if (!Status) {
      errors.Status = "Status is required.";
      isValid = false;
    }
    if (!Category) {
      errors.Category = "Category is required.";
      isValid = false;
    }
    if (!AssignedBy.Title) {
      errors.AssignedBy = "Assigned By is required.";
      isValid = false;
    }

    this.setState({ fieldErrors: errors });
    return isValid;
  };

  handleBlur = (
    event: React.FocusEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement | HTMLDivElement>,
    fieldName: keyof IAddTaskState
  ): void => {
    let error = "";
  
    // Check if the event target is an HTMLInputElement, HTMLTextAreaElement, or HTMLSelectElement
    if (
      event.target instanceof HTMLInputElement ||
      event.target instanceof HTMLTextAreaElement ||
      event.target instanceof HTMLSelectElement
    ) {
      if (!event.target.value.trim()) {
        error = `${fieldName} is required.`;
      }
    } else if (event.target instanceof HTMLDivElement) {
      
      const dropdownValue = this.state[fieldName] as string; 
  
      if (!dropdownValue) {
        error = `${fieldName} is required.`;
      }
    }
  
    this.setState((prevState) => ({
      ...prevState,
      fieldErrors: { ...prevState.fieldErrors, [fieldName]: error },
    }));
  };
  
  

  handleSubmit = async (): Promise<void> => {
    const { context } = this.props;
    const { Title, Description, DueDate, Priority, Status, Category, AssignedBy } = this.state;

    if (!this.validateForm()) {
      return;
    }
  
    const { value: confirmed } = await Swal.fire({
      title: 'Do you want to submit?',
      text: "Please confirm that you want to submit the form.",
      icon: 'question',
      showCancelButton: true,
      confirmButtonColor: '#3085d6',
      cancelButtonColor: '#d33',
      confirmButtonText: 'Yes, submit it!',
      cancelButtonText: 'Cancel',
    });
  
    if (!confirmed) {
      return; 
    }
  
    this.setState({ isSubmitting: true, errorMessage: "" });
  
    try {
      await TaskService.addTask(context, {
        Title, Description, DueDate, Priority, Status, Category, AssignedBy,
        Id: 0
      });
  
      this.setState({
        Title: "",
        Description: "",
        DueDate: "",
        Priority: {
          Id: "",
          Title: ""
        },
        Status: "",
        Category: "",
        AssignedBy: {
          Id: "",
          Title: "",
        },
        peoplePickerItems: [],
        isSubmitting: false,
        fieldErrors: {
          Title: "",
          Description: "",
          DueDate: "",
          Priority: "",
          Status: "",
          Category: "",
          AssignedBy: "",
        },
      });
  
      await Swal.fire("Success", "Task added successfully!", "success");
  
      const baseUrl = `https://indica.sharepoint.com/sites/TenantPracticeSite/SitePages/TaskDashbaord.aspx`;
      window.location.href = baseUrl;
  
    } catch (error) {
      this.setState({
        isSubmitting: false,
        errorMessage: error.message || "Failed to add the task. Please try again.",
      });
      await Swal.fire("Error", "Failed to add the task. Please try again.", "error");
    }
  };

  handleCancel = (): void => {
    this.setState({
      Title: "",
      Description: "",
      DueDate: "",
      Priority: {
        Id:"",
        Title: "",
      },
      Status:"",
      Category: "",
      AssignedBy: {
        Id:"",
        Title: "",
      },
      peoplePickerItems: [],
      isSubmitting: false,
      fieldErrors: {
        Title: "",
        Description: "",
        DueDate: "",
        Priority: "",
        Status: "",
        Category: "",
        AssignedBy: "",
      },
    });
  }

  renderToDashBoard = (): void => {
    const baseUrl = `https://indica.sharepoint.com/sites/TenantPracticeSite/SitePages/TaskDashbaord.aspx`;
      window.location.href = baseUrl;
  }

  render(): React.ReactNode {

    const { Title, Description, DueDate, Priority, Status, Category, isSubmitting, fieldErrors } = this.state;

    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient,
    };

    return (
      <div
        style={{
          backgroundImage:
            "url('https://images.unsplash.com/photo-1519681393784-d120267933ba?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1124&q=100')",
          backgroundPosition: "center",
          backgroundSize: "cover",
          height: "100%", 
          display: "flex",
          justifyContent: "center",
          alignItems: "center",
          flexDirection: "column",
        }}
      >
        <div className={styles.container}>
          <button className={styles.goToList} onClick={this.renderToDashBoard}>
              <span style={{
                paddingTop: "5px",
              }}
              >
                Go to DashBaord
              </span>
              <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 74 74"
                height="34"
                width="34"
              >
                <circle stroke-width="3" stroke="black" r="35.5" cy="37" cx="37"/>
                <path
                  fill="black"
                  d="M25 35.5C24.1716 35.5 23.5 36.1716 23.5 37C23.5 37.8284 24.1716 38.5 25 38.5V35.5ZM49.0607 38.0607C49.6464 37.4749 49.6464 36.5251 49.0607 35.9393L39.5147 26.3934C38.9289 25.8076 37.9792 25.8076 37.3934 26.3934C36.8076 26.9792 36.8076 27.9289 37.3934 28.5147L45.8787 37L37.3934 45.4853C36.8076 46.0711 36.8076 47.0208 37.3934 47.6066C37.9792 48.1924 38.9289 48.1924 39.5147 47.6066L49.0607 38.0607ZM25 38.5L48 38.5V35.5L25 35.5V38.5Z"
                />
              </svg>
          </button>
        </div>
        <Label
          styles={{
            root: {
              fontSize: "20px",
              fontWeight: 600,
              marginTop: "20px",
              marginBottom: "20px", 
              textAlign: "center",
              color: "white",
            },
          }}
        > 
          Add Task 
        </Label>
          
        <Stack
          tokens={{ childrenGap: 20 }}
          styles={{
            root: {
              width: "100%",
              maxHeight: "100vh",
              maxWidth: 600,
              margin: "20px 20px",
              padding: 40,
              borderRadius: "12px",
              backdropFilter: "blur(10px) saturate(180%)",
              WebkitBackdropFilter: "blur(16px) saturate(180%)",
              backgroundColor: "rgba(17, 25, 40, 0.75)",
              border: "1px solid rgba(255, 255, 255, 0.125)",
              boxShadow: "6px 6px 12px rgba(0, 0, 0, 0.2)",
              display: "flex",
              flexDirection: "column",
              overflowY: "auto", 
              flexShrink: 0,
            },
          }}
        >
          
          <TextField
            label="Task Title"
            name="Title"
            value={Title}
            placeholder="Enter the task title"
            onChange={this.handleInputChange}
            onBlur={(e) => this.handleBlur(e, "Title")}
            required
            errorMessage={fieldErrors.Title}
            styles={{
              fieldGroup: {
                backgroundColor: 'transparent', 
                borderColor: 'white',
                color: 'white',
              },
              field: {
                color: 'white',
                '::placeholder': { color: '#cccccc', opacity: 1 },
              },
              subComponentStyles: {
                label: { root: { color: 'white' } },
              },
              errorMessage: { color: '#ff8a8a' }, 
            }}
          />
    
          <TextField
            label="Task Description"
            name="Description"
            value={Description}
            multiline
            rows={3}
            placeholder="Provide a brief description of the task"
            onChange={this.handleInputChange}
            onBlur={(e) => this.handleBlur(e, "Description")}
            required
            errorMessage={fieldErrors.Description}
            styles={{
              fieldGroup: {
                borderColor: 'white',
                backgroundColor: 'transparent',
              },
              field: {
                color: 'white',
                '::placeholder': { color: '#cccccc', opacity: 1 },
              },
              subComponentStyles: {
                label: { root: { color: 'white' } },
              },
              errorMessage: { color: '#ff8a8a' }, 
            }}
          />

          <Stack horizontal tokens={{ childrenGap: 15 }}>
            <Stack.Item styles={{ root: { width: "50%" } }}>
              <Label
                styles={{
                  root: {
                    color: 'white',
                    marginBottom: '2px',
                    padding: '4px 0px',
                  },
                }}
              >
                Assigned By <span style={{color: '#AA2F33'}}>*</span>
              </Label>
              <PeoplePicker
                  context={peoplePickerContext}
                  placeholder="Enter assigned by name"
                  personSelectionLimit={1}
                  groupName={""}
                  showtooltip={true}
                  required
                  onChange={this.handlePeoplePickerChange}
                  principalTypes={[PrincipalType.User]}
                  ensureUser ={true}
                  styles={{
                      root: {
                          backgroundColor: 'transparent',
                          color: 'white',
                      },
                      text: {
                          borderColor: 'white',
                          color: 'white',
                      },
                      input: {
                          color: 'white',
                          backgroundColor: 'transparent',
                          '::placeholder': {
                              color: '#cccccc',
                              opacity: 1,
                          },
                      },
                  }}
              />
              {fieldErrors.AssignedBy && (
                <span
                  style={{
                    color: "#ff8a8a", 
                    fontSize: "12px",
                    marginTop: "5px",
                    display: "block",
                  }}
                >
                  {fieldErrors.AssignedBy}
                </span>
              )}
            </Stack.Item>

            <Stack.Item styles={{ root: { width: "50%" } }}>
              <TextField
                label="Due Date"
                name="DueDate"
                type="date"
                value={DueDate}
                onChange={this.handleInputChange}
                onBlur={(e) => this.handleBlur(e, "DueDate")}
                required
                errorMessage={fieldErrors.DueDate}
                styles={{
                  fieldGroup: {
                    backgroundColor: 'transparent',
                    borderColor: 'white',
                  },
                  field: {
                    color: "white",
                    "::placeholder": {
                      color: "#cccccc",
                    },
                  },
                  subComponentStyles: {
                    label: { root: { color: 'white' } },
                  },
                  errorMessage: { color: '#ff8a8a' }, 
                }}
              />
            </Stack.Item>
          </Stack>
          
          <Stack horizontal tokens={{ childrenGap: 15 }}>
            <Stack.Item styles={{ root: { width: "50%" } }}>
              <Dropdown
                label="Priority"
                options={[
                  { key: "High", text: "High" },
                  { key: "Medium", text: "Medium" },
                  { key: "Low", text: "Low" },
                ]}
                selectedKey={Priority.Title}
                data-field-name="Priority"
                onChange={this.handleDropdownChange}
                onBlur={(e) => this.handleBlur(e, "Priority")}
                required
                errorMessage={fieldErrors.Priority}
                placeholder="Select priority level"
                styles={{
                  subComponentStyles: {
                    label: { root: { color: 'white' } }, 
                  },
                  title: {
                    backgroundColor: 'transparent !important', 
                    color: Status ? 'white !important' : '#cccccc !important' ,
                    borderColor: 'white !important', 
                  },
                  dropdown: {
                    backgroundColor: 'transparent !important', 
                    color: 'white !important',
                  },
                  dropdownItem: {
                    backgroundColor: 'transparent !important', 
                    color: 'black !important', 
                    selectors: {
                      ':hover': {
                        backgroundColor: 'rgba(255, 255, 255, 0.1) !important', 
                        color: 'black !important', 
                      },
                    },
                  },
                  dropdownItemSelected: {
                    backgroundColor: 'rgba(255, 255, 255, 0.2) !important', 
                    color: Status ? 'black !important' : 'black !important' , 
                  },
                  errorMessage: {
                    color: '#ff8a8a !important',
                  },
                }}
              />
            </Stack.Item>

            <Stack.Item styles={{ root: { width: "50%" } }}> 
              <Dropdown
                label="Status"
                options={[
                  { key: "Pending", text: "Pending" },
                  { key: "Completed", text: "Completed" },
                  { key: "In Progress", text: "In Progress" },
                ]}
                selectedKey={Status}
                data-field-name="Status"
                onChange={this.handleDropdownChange}
                onBlur={(e) => this.handleBlur(e, "Status")}
                required
                errorMessage={fieldErrors.Status}
                placeholder="Select task status"
                styles={{
                  subComponentStyles: {
                    label: { root: { color: 'white' } }, 
                  },
                  title: {
                    backgroundColor: 'transparent !important', 
                    color: Status ? 'white !important' : '#cccccc !important' ,
                    borderColor: 'white !important', 
                  },
                  dropdown: {
                    backgroundColor: 'transparent !important', 
                    color: 'white !important',
                  },
                  dropdownItem: {
                    backgroundColor: 'transparent !important', 
                    color: 'black !important', 
                    selectors: {
                      ':hover': {
                        backgroundColor: 'rgba(255, 255, 255, 0.1) !important', 
                        color: 'black !important', 
                      },
                    },
                  },
                  dropdownItemSelected: {
                    backgroundColor: 'rgba(255, 255, 255, 0.2) !important', 
                    color: Status ? 'black !important' : 'black !important' , 
                  },
                  errorMessage: {
                    color: '#ff8a8a !important',
                  },
                }}
              />
            </Stack.Item>
          </Stack>

          <TextField
            label="Category"
            name="Category"
            value={Category}
            placeholder="E.g., Development, Design, Testing"
            onChange={this.handleInputChange}
            onBlur={(e) => this.handleBlur(e, "Category")}
            required
            errorMessage={fieldErrors.Category}
            styles={{
              fieldGroup: {
                backgroundColor: 'transparent',
                borderColor: 'white',
              },
              field: {
                color: 'white',
                '::placeholder': { color: '#cccccc', opacity: 1 },
              },
              subComponentStyles: {
                label: { root: { color: 'white' } },
              },
              errorMessage: { color: '#ff8a8a' }, 
            }}
          />
    
          <Stack horizontal tokens={{ childrenGap: 10 }} horizontalAlign="center">
            <PrimaryButton
              text={isSubmitting ? "submitting..." : "Add Task"}
              disabled={isSubmitting}
              onClick={this.handleSubmit}
              styles={{
                root: {
                  width: "150px",
                  color: "#090909",  
                  padding: "0.7em 1.7em",  
                  fontSize: "15px", 
                  borderRadius: "0.5em", 
                  backgroundColor: "#e8e8e8", 
                  cursor: "pointer", 
                  border: "1px solid #e8e8e8", 
                  transition: "all 0.3s",
                },
              }}
            />
            <DefaultButton
              text="Cancel"
              onClick={this.handleCancel}
              styles={{
                root: {
                  width: "150px",
                  color: "#090909",
                  padding: "0.7em 1.7em", 
                  fontSize: "15px", 
                  borderRadius: "0.5em", 
                  backgroundColor: "#e8e8e8",
                  cursor: "pointer",
                  border: "1px solid #e8e8e8", 
                  transition: "all 0.3s", 
                },
              }}
            />
          </Stack>
        </Stack>
      </div>
    );
  }
}
