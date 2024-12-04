import { sp } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITaskItem {
  Id: number;
  Title: string;
  Description: string;
  DueDate: string;
  Priority: {
    Id: string;
    Title: string;
  } | undefined;
  Status: string;
  AssignedBy: {
    Id?: string;
    Title?: string;
  } | undefined;
  Category: string;
}

export interface IEditTaskItem {
  ID: number;
  Title?: string;
  Description?: string;
  DueDate?: string;
  Priority?: {
    Id: string;
    Title: string;
  };
  Status?: string;
  AssignedBy?: {
    Id?: string;
    Title?: string;
  };
  Category?: string;
  Comment?: string;
}

const taskListName = "TaskManagerList";
const metadataListName = "TaskManagerMetadata";

const convertToDisplayDate = (isoDateString: string): string => {
  if (!isoDateString) return "";

  const date = new Date(isoDateString);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`; 
};


export class TaskService {
  public static async getTasks(context: WebPartContext): Promise<ITaskItem[]> {
    sp.setup({
      spfxContext: {
        ...context,
        pageContext: context.pageContext,
        msGraphClientFactory: {
          getClient: async () => await context.msGraphClientFactory.getClient("3"),
        },
      },
    });

    try {
      const items = await sp.web.lists
        .getByTitle(taskListName)
        .items.select(
          "Id",
          "Title",
          "Description",
          "DueDate",
          "Priority/Id",
          "Priority/Title",
          "Status",
          "AssignedBy/Id",
          "AssignedBy/Title",
          "Category"
        )
        .expand("Priority,AssignedBy")
        .get();

      return items.map((item) => ({
        Id: item.Id,
        Title: item.Title,
        Description: item.Description,
        DueDate: item.DueDate ? new Date(item.DueDate).toLocaleDateString() : "",
        Priority: item.Priority
          ? { Id: item.Priority.Id, Title: item.Priority.Title }
          : undefined,
        Status: item.Status,
        AssignedBy: item.AssignedBy
          ? { Id: item.AssignedBy.Id, Title: item.AssignedBy.Title }
          : undefined,
        Category: item.Category || "Uncategorized",
      }));
    } catch (error) {
      console.error("Error fetching tasks:", error);
      throw new Error("Failed to fetch tasks");
    }
  }

  public static async getTaskById(context: WebPartContext, taskId: number): Promise<ITaskItem | undefined> {
    sp.setup({
      spfxContext: {
        ...context,
        pageContext: context.pageContext,
        msGraphClientFactory: {
          getClient: async () => await context.msGraphClientFactory.getClient("3"),
        },
      },
    });
    console.log(taskId);
    

    try {
      const task = await sp.web.lists
        .getByTitle(taskListName)
        .items.getById(taskId)
        .select(
          "Id",
          "Title",
          "Description",
          "DueDate",
          "Priority/Id",
          "Priority/Title",
          "Status",
          "AssignedBy/Id",
          "AssignedBy/Title",
          "Category"
        )
        .expand("Priority,AssignedBy")
        .get();
        console.log("fetched edited task",task);
        
      return {
        Id: task.Id,
        Title: task.Title,
        Description: task.Description || "",
        DueDate:  convertToDisplayDate(task.DueDate) || "",
        Priority: task.Priority
          ? { Id: task.Priority.Id, Title: task.Priority.Title }
          : undefined,
        Status: task.Status,
        AssignedBy: task.AssignedBy
          ? { Id: task.AssignedBy.Id, Title: task.AssignedBy.Title }
          : undefined,
        Category: task.Category || "",
      };
    } catch (error) {
      console.error(`Error fetching task with ID ${taskId}:`, error);
      return undefined;
    }
  }

  public static async addTask(context: WebPartContext, taskData: ITaskItem): Promise<void> {
    sp.setup({
        spfxContext: {
            ...context,
            pageContext: context.pageContext,
            msGraphClientFactory: {
                getClient: async () => await context.msGraphClientFactory.getClient('3'),
            },
        },
    });

    const { Title, Description, DueDate, Priority, Status, Category, AssignedBy } = taskData;

    try {
        const assignedByEmail = AssignedBy ? AssignedBy.Title : '';
        if (!assignedByEmail) {
            throw new Error("AssignedBy email is required.");
        }

        const ensureUserResult = await sp.web.ensureUser(assignedByEmail);
        const assignedById = ensureUserResult.data.Id;

        const priorityItems = await sp.web.lists
            .getByTitle(metadataListName) 
            .items.filter(`Title eq '${Priority?.Title}'`)
            .get();

        if (priorityItems.length === 0) {
            throw new Error(`Priority '${Priority}' not found in metadata list.`);
        }
        const priorityId = priorityItems[0].Id;

        await sp.web.lists.getByTitle(taskListName).items.add({
            Title,
            Description,
            DueDate,
            PriorityId: priorityId, 
            Status,    
            Category,
            AssignedById: assignedById,
        });

        console.log("Task added successfully.");
    } catch (error) {
        console.error("Error adding task:", error);
        throw new Error("Failed to add task");
    }
  }

  public static async editTask(context: WebPartContext, taskItem: IEditTaskItem): Promise<void> {
    sp.setup({
      spfxContext: {
        ...context,
        pageContext: context.pageContext,
        msGraphClientFactory: {
          getClient: async () => await context.msGraphClientFactory.getClient("3"),
        },
      },
    });

    try {

      const formattedDueDate = taskItem.DueDate
        ? new Date(taskItem.DueDate).toISOString().split(".")[0] + "Z" 
        : undefined;

      const updateData = Object.fromEntries(
        Object.entries({
          Title: taskItem.Title,
          Description: taskItem.Description,
          DueDate: formattedDueDate,
          PriorityId: taskItem.Priority?.Id,
          Status: taskItem.Status,
          AssignedById: taskItem.AssignedBy?.Id,
          Category: taskItem.Category,
          Comment: taskItem.Comment,
        }).filter(([_, value]) => value !== undefined && value !== null)
      );

      if (Object.keys(updateData).length === 0) {
        throw new Error("No data to update");
      }

      console.log("Update task object", updateData);
      

      await sp.web.lists.getByTitle(taskListName).items.getById(taskItem.ID).update(updateData);
    } catch (error) {
      console.error("Error updating task:", error);
      throw new Error("Failed to update task");
    }
  }

  public static sortTasks(
    tasks: ITaskItem[],
    columnKey: keyof ITaskItem,
    isDescending: boolean
  ): ITaskItem[] {
    return [...tasks].sort((a, b) => {
      let valueA: string | undefined;
      let valueB: string | undefined;
  
      if (columnKey === 'Priority') {
        valueA = a.Priority?.Title; 
        valueB = b.Priority?.Title;
      } else if (columnKey === 'AssignedBy') {
        valueA = a.AssignedBy?.Title;
        valueB = b.AssignedBy?.Title;
      } else {
        valueA = a[columnKey] as string | undefined; 
        valueB = b[columnKey] as string | undefined;
      }
  
      if (valueA === undefined || valueB === undefined) {
        return valueA === valueB ? 0 : valueA === undefined ? 1 : -1;
      }
  
      const finalA = String(valueA);
      const finalB = String(valueB);
  
      if (finalA < finalB) return isDescending ? 1 : -1;
      if (finalA > finalB) return isDescending ? -1 : 1;
      return 0;
    });
  }


  public static async deleteTask(taskId: number): Promise<void> {
    try {
      await sp.web.lists.getByTitle(taskListName).items.getById(taskId).delete();
    } catch (error) {
      console.error(`Error deleting task with ID ${taskId}:`, error);
      throw new Error("Failed to delete task");
    }
  }
}
