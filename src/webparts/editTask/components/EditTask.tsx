import * as React from "react";
import { Stack, TextField, Dropdown, IDropdownOption, PrimaryButton, Label, DefaultButton, IPersonaProps } from "@fluentui/react";
import { IPeoplePickerContext, PeoplePicker, PrincipalType} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import Swal from "sweetalert2";
import { IEditTaskProps } from "./IEditTaskProps";
import { IEditTaskItem, TaskService } from "../../../Services/TaskService";

interface IEditTaskState {
  ID: number;
  Title: string;
  Description: string;
  DueDate: string;
  Priority: {
    Id: string,
    Title: string,
  };
  Status: string;
  Category: string;
  AssignedBy: {
    Id?: string,
    Title?: string,
  };
  Comment: string;
  isSubmitting: boolean;
  errorMessage: string;
  fieldErrors: { [key in keyof IEditTaskState]?: string };
}

const params = new URLSearchParams(window.location.search);
const taskId = params.get("taskId");

if (!taskId || isNaN(Number(taskId))) {
  console.error("Invalid task ID in URL:", taskId);
}
console.log("Valid task ID:", taskId);


export default class EditTask extends React.Component<IEditTaskProps, IEditTaskState> {
  constructor(props: IEditTaskProps) {
    super(props);
    this.state = {
      ID: 0,
      Title: "",
      Description: "",
      DueDate: "",
      Priority: {
        Id: "",
        Title: "",
      },
      Status:"",
      Category: "",
      AssignedBy: {
        Id: "",
        Title: "",
      },
      Comment: "",
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
        Comment: "",
      },
    };
  }

  async componentDidMount():Promise<void> {
    if (taskId) {
      await this.fetchTaskDetails(Number(taskId));
    }
  }

  fetchTaskDetails = async (taskId: number): Promise<void> => {
    console.log("Fetching details for task ID:", taskId);
  
    try {
        const taskData = await TaskService.getTaskById(this.props.context, taskId);
  
        if (taskData) {
            console.log("Fetched Task Data by id:", taskData);
  
            this.setState({
                ID: taskData.Id,
                Title: taskData.Title,
                Description: taskData.Description,
                DueDate: taskData.DueDate,
                Priority: taskData.Priority
                  ? { Id: taskData.Priority.Id, Title: taskData.Priority.Title }
                  : { Id: "", Title: "" },
                Status: taskData.Status,
                Category: taskData.Category,
                AssignedBy: taskData.AssignedBy
                  ? { Id: taskData.AssignedBy.Id, Title: taskData.AssignedBy.Title }
                  : { Id: "", Title: "" }, 
            });
        } else {
            console.error("Task data not found for ID:", taskId);
        }
    } catch (error) {
        console.error("Failed to fetch task details:", error);
    }
  };

  formatDate = (dateString: string): string => {
    if (!dateString) {
        return ""; 
    }

    const date = new Date(dateString);

    if (isNaN(date.getTime())) {
        console.error("Invalid date string:", dateString);
        return "";
    }

    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear(); 

    return `${day}-${month}-${year}`;
  };
  

  handlePeoplePickerChange = (items: IPersonaProps[]): void => {
    console.log("PeoplePicker selection:", items);
  
    if (items.length > 0) {
      const selectedUser = items[0];
      console.log("Selected User:", selectedUser);
  
      this.setState({
        AssignedBy: { Id: selectedUser.id || "", Title: selectedUser.text || "" },
        fieldErrors: { ...this.state.fieldErrors, AssignedBy: "" },
      });
    } else {
      this.setState({
        AssignedBy: { Id: "", Title: "" },
        fieldErrors: { ...this.state.fieldErrors, AssignedBy: "Assigned By is required." },
      });
    }
  };
  
  handleInputChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    const { name } = event.currentTarget;
    this.setState({
      [name]: newValue || "",
      fieldErrors: { ...this.state.fieldErrors, [name]: "" },
    } as unknown as Pick<IEditTaskState, keyof IEditTaskState>);
  };

  
  convertToDisplayDate = (isoDateString: string): string => {
    if (!isoDateString) return "";
  
    const date = new Date(isoDateString);
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
  
    return `${year}-${month}-${day}`; 
  };
  
  
  convertToISODateFormat = (displayDateString: string): string => {
    if (!displayDateString) return "";
  
    const [day, month, year] = displayDateString.split("-");
    return `${year}-${month}-${day}T00:00:00Z`;
  };

  handleDateInputChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>
  ): void => {
    const { name, value } = event.currentTarget;
  
    this.setState({
      [name]: value, 
      fieldErrors: { ...this.state.fieldErrors, [name]: "" },
    } as unknown as Pick<IEditTaskState, keyof IEditTaskState>);
  };
  

  handleDropdownChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    const { key, text } = option || {};
    const fieldName = (event.target as HTMLDivElement).getAttribute("data-field-name");
  
    if (fieldName) {
      this.setState((prevState) => {
        let updatedField: string | { Id: string, Title: string };
  
        if (fieldName === "Priority") {
          updatedField = { Id: key as string, Title: text || "" }; 
        } else {
          updatedField = key as string;
        }
  
        return {
          ...prevState,
          [fieldName]: updatedField, 
          fieldErrors: {
            ...prevState.fieldErrors,
            [fieldName]: key ? "" : "Please select a value",
          },
        };
      });
    }
  };
  
  validateForm = (): boolean => {
    const { Title, Description, DueDate, Priority, Status, Category, AssignedBy, Comment } = this.state;
    let isValid = true;
    const errors = { ...this.state.fieldErrors };

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
    if (!Comment.trim()) {
      errors.Comment = "Comment is required.";
      isValid = false;
    }

    this.setState({ fieldErrors: errors });
    return isValid;
  };


  handleSubmit = async (): Promise<void> => {
    const { context } = this.props;
    const { Title, Description, DueDate, Priority, Status, Category, AssignedBy, Comment, ID } = this.state;
  
    if (!this.validateForm()) {
      return;
    }
  
    this.setState({ isSubmitting: true, errorMessage: "" });
  
    try {

      const taskItem: IEditTaskItem = {
        Title,
        Description,
        DueDate,
        Priority: Priority ? { Id: Priority.Id, Title: Priority.Title } : undefined, 
        Status,
        Category,
        AssignedBy,
        Comment,
        ID,
      };

      await TaskService.editTask(context, taskItem);
  
      this.setState({
        Title: "",
        Description: "",
        DueDate: "",
        Priority: {
          Id: "",
          Title: "",
        },
        Status: "",
        Category: "",
        AssignedBy: {
          Id: "",
          Title: "",
        },
        Comment: "",
        isSubmitting: false,
        fieldErrors: {
          Title: "",
          Description: "",
          DueDate: "",
          Priority: "",
          Status: "",
          Category: "",
          AssignedBy: "",
          Comment: "",
        },
      });
  
      await Swal.fire("Success", "Task updated successfully!", "success");
      const baseUrl = `https://indica.sharepoint.com/sites/TenantPracticeSite/SitePages/TaskDashbaord.aspx`;
      window.location.href = baseUrl;
    } catch (error) {
      this.setState({
        isSubmitting: false,
        errorMessage: error.message || "Failed to update the task. Please try again.",
      });
  
      await Swal.fire("Error", error.message || "Failed to update the task. Please try again.", "error");
    }
  };

  handleCancel = (): void => {
    const baseUrl = `https://indica.sharepoint.com/sites/TenantPracticeSite/SitePages/TaskDashbaord.aspx`;
      window.location.href = baseUrl;
  }
  

  render(): React.ReactNode {
    const { Title, Description, AssignedBy, DueDate, Priority, Status, Category, Comment, isSubmitting, fieldErrors } = this.state;

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
          Update Task 
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
            onChange={this.handleInputChange}
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
            onChange={this.handleInputChange}
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
                Assigned By
              </Label>
              <PeoplePicker
                context={peoplePickerContext}
                placeholder="Enter assigned by name"
                personSelectionLimit={1}
                onChange={this.handlePeoplePickerChange}
                principalTypes={[PrincipalType.User]}
                ensureUser={true}
                defaultSelectedUsers={AssignedBy?.Title ? [AssignedBy?.Title] : []} 
                styles={{
                  root: {
                    backgroundColor: 'transparent',
                  },
                  text: {
                    borderColor: 'white',
                    color: 'white',
                  },
                  input: {
                    color: 'white',
                    backgroundColor: 'transparent',
                    '::placeholder': { color: '#cccccc', opacity: 1 },
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
                value={this.convertToDisplayDate(DueDate)} 
                onChange={this.handleDateInputChange}
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
                  { key: 3, text: "High" },
                  { key: 2, text: "Medium" },
                  { key: 1, text: "Low" },
                ]}
                selectedKey={Priority.Id}
                onChange={this.handleDropdownChange}
                errorMessage={fieldErrors.Priority}
                data-field-name="Priority"
                styles={{
                  subComponentStyles: {
                    label: { root: { color: 'white' } }, 
                  },
                  title: {
                    backgroundColor: 'transparent !important', 
                    color: Priority ? 'white !important' : '#cccccc !important' ,
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
                    color: Priority ? 'black !important' : 'black !important' , 
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
                selectedKey={Status || ""}
                onChange={this.handleDropdownChange}
                data-field-name="Status"
                errorMessage={fieldErrors.Status}
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
                    color: Priority ? 'black !important' : 'black !important' , 
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
            onChange={this.handleInputChange}
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

          <TextField
            label="Comment"
            name="Comment"
            value={Comment}
            placeholder="Enter your comment"
            onChange={this.handleInputChange}
            required
            errorMessage={fieldErrors.Comment}
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
              text={isSubmitting ? "Updateting..." : "Update Task"}
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
                rootPressed: {
                  color: "#666",
                }
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
                rootPressed: {
                  color: "#666",
                }
              }}
            />
          </Stack>
        </Stack>
      </div>
    );
  
  }
}
