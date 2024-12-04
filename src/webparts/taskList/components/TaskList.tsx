import * as React from "react";
import { ITaskListProps } from "./ITaskListProps";
import { TaskService, ITaskItem } from "../../../Services/TaskService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  DetailsList,
  IColumn,
  Spinner,
  MessageBar,
  MessageBarType,
  SearchBox,
  Label,
  IconButton,
} from "@fluentui/react";
import styles from "./TaskList.module.scss";

interface ITaskDashboardState {
  tasks: ITaskItem[];
  filteredTasks: ITaskItem[];
  isLoading: boolean;
  error: string | undefined;
  searchQuery: string;
  sortKey: keyof ITaskItem | undefined;
  isDescending: boolean;
  currentPage: number;
  taskPerPage: number;
}

export default class TaskList extends React.Component<
  ITaskListProps,
  ITaskDashboardState
> {
  constructor(props: ITaskListProps) {
    super(props);
    this.state = {
      tasks: [],
      filteredTasks: [],
      isLoading: false,
      error: undefined,
      searchQuery: "",
      sortKey: undefined,
      isDescending: false,
      currentPage: 1,
      taskPerPage: 5,
    };
  }

  async componentDidMount(): Promise<void> {
    await this.fetchTasks();
  }

  private fetchTasks = async (): Promise<void> => {
    this.setState({ isLoading: true, error: undefined });
    try {
      const tasks = await TaskService.getTasks(this.props.context as WebPartContext);
      this.setState({ tasks, filteredTasks: tasks, isLoading: false });
    } catch (error) {
      this.setState({
        error: error.message || "Failed to load tasks",
        isLoading: false,
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
          || (task.Priority?.Title?.toLowerCase().includes(query) ?? false)
      );
    });

    this.setState({ searchQuery: query, filteredTasks });
  };

  private handleSort = (columnKey: keyof ITaskItem): void => {
    const { filteredTasks, isDescending, sortKey } = this.state;
    const newIsDescending = columnKey === sortKey ? !isDescending : false;

    const sortedTasks = TaskService.sortTasks(
      filteredTasks,
      columnKey,
      newIsDescending
    );

    this.setState({
      filteredTasks: sortedTasks,
      sortKey: columnKey,
      isDescending: newIsDescending,
    });
  };

  private changePage = (page: number): void => {
    this.setState({ currentPage: page });
  };

  private handlePageInputChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const page = Number(event.target.value);
    if (page > 0 && page <= Math.ceil(this.state.filteredTasks.length / this.state.taskPerPage)) {
      this.changePage(page);
    }
  };

  private getCurrentPageTasks(): ITaskItem[] {
    const { filteredTasks, currentPage, taskPerPage } = this.state;
    const startIndex = (currentPage - 1) * taskPerPage;
    return filteredTasks.slice(startIndex, startIndex + taskPerPage);
  }

  private getPaginationControls(): React.ReactNode {
    const { filteredTasks, currentPage, taskPerPage } = this.state;
    const totalPages = Math.ceil(filteredTasks.length / taskPerPage);

    return (
      <div className={styles.pagination}>
        <IconButton
          className={styles.iconButton}
          iconProps={{ iconName: "ChevronLeft" }}
          disabled={currentPage === 1}
          onClick={() => this.changePage(currentPage - 1)}
        />
        {Array.from({ length: totalPages }, (_, i) => (
          <span
            key={i + 1}
            className={`${styles.pageNumber} ${
              currentPage === i + 1 ? "active" : ""
            }`}
            onClick={() => this.changePage(i + 1)}
          >
            {i + 1}
          </span>
        ))}
        <IconButton
          className={styles.iconButton}
          iconProps={{ iconName: "ChevronRight" }}
          disabled={currentPage === totalPages}
          onClick={() => this.changePage(currentPage + 1)}
        />
        <input
          type="number"
          min="1"
          max={totalPages}
          value={currentPage}
          onChange={this.handlePageInputChange}
          className={styles.pageInput}
          style={{ width: '50px', marginLeft: '10px' }}
        />
      </div>
    );
  }

  private getColumns(): IColumn[] {
    const { sortKey, isDescending } = this.state;

    return [
      {
        key: "Title",
        name: "Title",
        fieldName: "Title",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        isSorted: sortKey === "Title",
        isSortedDescending: isDescending,
        onColumnClick: () => this.handleSort("Title"),
      },
      {
        key: "Description",
        name: "Description",
        fieldName: "Description",
        isMultiline: true,
        minWidth: 150,
        maxWidth: 300,
        isResizable: true,
        isSorted: sortKey === "Description",
        isSortedDescending: isDescending,
        onColumnClick: () => this.handleSort("Description"),
      },
      {
        key: "DueDate",
        name: "Due Date",
        fieldName: "DueDate",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        isSorted: sortKey === "DueDate",
        isSortedDescending: isDescending,
        onColumnClick: () => this.handleSort("DueDate"),
      },
      {
        key: "Priority",
        name: "Priority",
        fieldName: "Priority",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        isSorted: sortKey === "Priority",
        isSortedDescending: isDescending,
        onColumnClick: () => this.handleSort("Priority"),
        onRender: (item: ITaskItem) => item.Priority?.Title || "No Priority",
      },
      {
        key: "Status",
        name: "Status",
        fieldName: "Status",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        isSorted: sortKey === "Status",
        isSortedDescending: isDescending,
        onColumnClick: () => this.handleSort("Status"),
        onRender: (item: ITaskItem) => item.Status || "No Status",
      },
      {
        key: "AssignedBy",
        name: "Assigned By",
        fieldName: "AssignedBy",
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        isSorted: sortKey === "AssignedBy",
        isSortedDescending: isDescending,
        onColumnClick: () => this.handleSort("AssignedBy"),
        onRender: (item: ITaskItem) => item.AssignedBy?.Title || "Unassigned",
      },
      {
        key: "Category",
        name: "Category",
        fieldName: "Category",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        isSorted: sortKey === "Category",
        isSortedDescending: isDescending,
        onColumnClick: () => this.handleSort("Category"),
      },
    ];
  }

  renderToDashBoard = (): void => {
    const baseUrl = `https://indica.sharepoint.com/sites/TenantPracticeSite/SitePages/TaskDashbaord.aspx`;
    window.location.href = baseUrl;
  }

  renderToAddTask = (): void => {
    const baseUrl = `https://indica.sharepoint.com/sites/TenantPracticeSite/SitePages/AddTask.aspx`;
    window.location.href = baseUrl;
  }

  public render(): React.ReactNode {
    const { isLoading, error, searchQuery } = this.state;
    const currentTasks = this.getCurrentPageTasks();

    return (
      <div
        style={{
          backgroundImage:
            "url('https://images.unsplash.com/photo-1519681393784-d120267933ba?ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&ixlib=rb-1.2.1&auto=format&fit=crop&w=1124&q=100')",
          backgroundPosition: "center",
          backgroundSize: "cover",
          width: "100 %",
          height: "100vh",
          display: "flex",
          justifyContent: "center",
          alignItems: "center",
          flexDirection: "column",
          padding: "20px",
          boxSizing: "border-box",
        }}
      >
        <Label
          styles={{
            root: {
              fontSize: "20px",
              fontWeight: 600,
              marginBottom: "20px",
              textAlign: "center",
              color: "white",
            },
          }}
        >
          Task List
        </Label>

        <div className={styles.container}>
          <button className={styles.goToList} onClick={this.renderToDashBoard}>
            <span style={{ paddingTop: "5px" }}>
              Go to Dashboard
            </span>
            <svg
              xmlns="http://www.w3.org/2000/svg"
              fill="none"
              viewBox="0 0 74 74"
              height="34"
              width="34"
            >
              <circle strokeWidth="3" stroke="black" r="35.5" cy="37" cx="37" />
              <path
                fill="black"
                d="M25 35.5C24.1716 35.5 23.5 36.1716 23.5 37C23.5 37.8284 24.1716 38.5 25 38.5V35.5ZM49.0607 38.0607C49.6464 37.4749 49.6464 36.5251 49.0607 35.9393L39.5147 26.3934C38.9289 25.8076 37.9792 25.8076 37.3934 26.3934C36.8076 26.9792 36.8076 27.9289 37.3934 28.5147L45.8787 37L37.3934 45.4853C36.8076 46.0711 36.8076 47.0208 37.3934 47.6066C37.9792 48.1924 38.9289 48.1924 39.5147 47.6066L49.0607 38.0607ZM25 38.5L48 38.5V35.5L25 35.5V38.5Z"
              />
            </svg>
          </button>

          <SearchBox
            placeholder="Search here..."
            value={searchQuery}
            onChange={this.handleSearch}
            styles={{
              root: {
                marginBottom: "15px",
                width: "100%",
                maxWidth: "250px",
                backgroundColor: 'rgba(255, 255, 255, 0.8)',
                borderRadius: '5px',
                alignItems: 'center',
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
              <path d="M0 0h24v24H0z" fill="none" />
              <path d="M11 11V5h2v6h6v2h-6v6h-2v-6H5v-2z" fill="currentColor" />
            </svg>
            <span>Add Task</span>
          </button>
        </div>

        {isLoading && <Spinner label="Loading tasks..." />}

        {error && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
          >
            {error}
          </MessageBar>
        )}

        {!isLoading && !error && currentTasks.length === 0 && (
          <MessageBar
            messageBarType={MessageBarType.warning}
            isMultiline={false}
          >
            No tasks available.
          </MessageBar>
        )}

        {!isLoading && currentTasks.length > 0 && (
          <div style={{ width: "100%", maxHeight: "100vh", marginTop: "20px" }}>
            <DetailsList
              items={currentTasks}
              columns={this.getColumns()}
              setKey="set"
              layoutMode={1}
              selectionMode={0}
              styles={{
                root: {
                  
                  width: "100%",
                  borderRadius: "12px",
                  backdropFilter: "blur(10px) saturate( 180%)",
                  WebkitBackdropFilter: "blur(16px) saturate(180%)",
                  backgroundColor: "transparent",
                  border: "1px solid rgba(255, 255, 255, 0.125)",
                  boxShadow: "6px 6px 12px rgba(0, 0, 0, 0.2)",
                  display: "flex",
                  flexDirection: "column",
                  boxSizing: "border-box",
                },
              }}
              isHeaderVisible={true}
            />
            {this.getPaginationControls()}
          </div>
        )}
      </div>
    );
  }
}