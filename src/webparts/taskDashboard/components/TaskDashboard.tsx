import * as React from "react";
import styles from "../../taskDashboard/components/TaskDashboard.module.scss"
import { TaskService,ITaskItem } from '../../../Services/TaskService';
import { ITaskDashboardProps } from "./ITaskDashboardProps";
import { SearchBox, Spinner } from "@fluentui/react";
import Swal from "sweetalert2";

export interface ITaskDashboardState {
  tasks: ITaskItem[];
  filteredTasks: ITaskItem[];
  isLoading: boolean;
  error?: string;
  searchQuery: string;
  sortKey?: keyof ITaskItem;
  isDescending: boolean;
}

export default class TaskDashboard extends React.Component<ITaskDashboardProps, ITaskDashboardState> {
  constructor(props: ITaskDashboardProps) {
    super(props);

    this.state = {
      tasks: [],
      filteredTasks: [],
      isLoading: true,
      error: undefined,
      searchQuery: "",
      sortKey: undefined,
      isDescending: false,
    };
  }

  async componentDidMount(): Promise<void> {
    await this.fetchTasks();
  }

  fetchTasks = async ():Promise<void> => {
    try {
      this.setState({ isLoading: true });

      const tasks = await TaskService.getTasks(this.props.context); 

      this.setState({
        tasks,
        filteredTasks: tasks,
        isLoading: false,
      });
    } catch (error) {
      this.setState({
        isLoading: false,
        error: "Failed to fetch tasks.",
      });
    }
  };

  private handleSearch = (
    event?: React.FormEvent<HTMLInputElement>,
    newValue?: string
  ): void => {
    const query = newValue?.toLowerCase() || "";
    const { tasks } = this.state;

    const filteredTasks = tasks.filter((task) => {
      return (
        Object.values(task).some(
          (value) =>
            typeof value === "string" && value.toLowerCase().includes(query)
        ) || (task.AssignedBy?.Title?.toLowerCase().includes(query) ?? false) 
          || (task.Priority?.Title?.toLowerCase().includes(query)?? false) 
      );
    });

    this.setState({ searchQuery: query, filteredTasks });
  };

  handleEditTask = (taskId: number): void => {
    const editTaskUrl = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/EditTask.aspx?taskId=${taskId}`;
    window.location.href = editTaskUrl;
  };
  
  
  handleDeleteTask = async (taskId: number): Promise<void> => {
    const { value: confirmed } = await Swal.fire({
      title: 'Are you sure?',
      text: "You won't be able to revert this!",
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#3085d6',
      cancelButtonColor: '#d33',
      confirmButtonText: 'Yes, delete it!',
      cancelButtonText: 'Cancel',
    });
  
    if (confirmed) {
      try {
        await TaskService.deleteTask(taskId);
        await this.fetchTasks();
        await Swal.fire('Deleted!', 'Your task has been deleted.', 'success');
      } catch (error) {
        console.error("Failed to delete task:", error);
        this.setState({ error: "Failed to delete task. Please try again." });
        await Swal.fire('Error!', 'Failed to delete task. Please try again.', 'error');
      }
    }
  };

  renderToList = (): void => {
    const baseUrl = `https://indica.sharepoint.com/sites/TenantPracticeSite/SitePages/TaskListTable.aspx`;
      window.location.href = baseUrl;
  }

  renderToAddTask = (): void => {
    const baseUrl = `https://indica.sharepoint.com/sites/TenantPracticeSite/SitePages/AddTask.aspx`;
      window.location.href = baseUrl;
  }
  
  render(): React.ReactNode {
    const { filteredTasks, isLoading, error, searchQuery } = this.state;

    return (
      <div className={styles.dashboard}>
        <button className={styles.goToList} onClick={this.renderToList}>
          <span>Go to List</span>
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

        <header>
          <h1 className={styles.title}>Task Dashboard</h1>
        </header>

        <header className={styles.header}>
          <SearchBox
            placeholder="Search here..."
            value={searchQuery}
            onChange={this.handleSearch}
            styles={{
              root: {
                marginTop: "5px",
                width: "100%",
                maxWidth: "250px",
                backgroundColor: 'transparent',
                borderRadius: '5px',
              },
            }}
          />

          <button className={styles.addButton} onClick={this.renderToAddTask}>
            <svg
              height="24"
              width="24"
              viewBox="0 0 24 24"
              xmlns="http://www.w3.org/2000/svg"
            >
              <path d="M0 0h24v24H0z" fill="none"/>
              <path d="M11 11V5h2v6h6v2h-6v6h-2v-6H5v-2z" fill="currentColor"/>
            </svg>
            <span>Add Task</span>
          </button>
        </header>

        {isLoading ? (
          //<p className={styles.loading}>Loading tasks...</p>
          <Spinner />
        ) : error ? (
          <p className={styles.error}>{error}</p>
        ) : filteredTasks.length === 0 ? (
          <p className={styles.noTasks}>No matching task data found.</p>
        ) : (
          <div className={styles.grid}>
            {filteredTasks.map((task, index) => (
              
              <div key={index} className={styles.card}>
                <div className={styles.cardHeader}>
                  <h3 className={styles.cardTitle}>{task.Title}</h3>
                  <p className={styles.cardDescription} >
                    {task.Description || "No description available."}
                  </p>
                </div>
                <div className={styles.cardContent}>
                  <dl className={styles.detailsList}>
                    <div className={styles.detailItem}>
                      <dt className={styles.detailKey}>Due Date</dt>
                      <dd className={styles.detailValue}>
                        {task.DueDate || "N/A"}
                      </dd>
                    </div>
                    <div className={styles.detailItem}>
                      <dt className={styles.detailKey}>Priority</dt>
                      <dd
                        className={styles.detailValue}
                        style={{
                          width:"30%",
                          textAlign: "right",
                          fontWeight: "bold",
                          backgroundColor:
                            task.Priority?.Title === "High"
                              ? "red"
                              : task.Priority?.Title === "Medium"
                              ? "yellow"
                              : "#90EE90",
                        }}
                      >
                        {task.Priority?.Title}
                      </dd>
                    </div>
                    <div className={styles.detailItem}>
                      <dt className={styles.detailKey}>Status</dt>
                      <dd className={styles.detailValue}>{task.Status}</dd>
                    </div>
                    <div className={styles.detailItem}>
                      <dt className={styles.detailKey}>Category</dt>
                      <dd className={styles.detailValue}>{task.Category}</dd>
                    </div>
                    <div className={styles.detailItem}>
                      <dt className={styles.detailKey}>Assigned By</dt>
                      <dd className={styles.detailValue}>
                        {task.AssignedBy && task.AssignedBy.Title || "Unassigned"}
                      </dd>
                    </div>
                  </dl>
                  <div className={styles.cardActions}>
                    <button
                      className={styles.editButton}
                      onClick={() => this.handleEditTask(task.Id)}
                    >
                      Edit
                    </button>
                    <button
                      className={styles.deleteButton}
                      onClick={() => this.handleDeleteTask(task.Id)}
                    >
                      Delete
                    </button>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }
}
