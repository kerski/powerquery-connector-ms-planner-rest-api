/*
    Author: John Kerski

    Description: Connector wraps many of the "Get" endpoints in the Graph API to access Microsoft Planner.  
    This connector serves as a way to have a library of Power Query functions to build datasets without the need
    for storing client secrets or passwords in the dataset.  The custom data connector will need
    to be installed in the a Power BI Gateway in order to refresh datasets leveraging this
    custom connector.  For more information on installing a custom data connector with a gateway
    please see: https://learn.microsoft.com/en-us/power-bi/connect-data/service-gateway-custom-connectors .
*/
[Version = "1.0.0"]
section MSPLANNERRESTAPI;
// Pull Client ID from Source File
client_id = Text.FromBinary(Extension.Contents("client_id.user")); 
redirect_uri = "https://oauth.powerbi.com/views/oauthredirect.html";
token_uri = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token";
authorize_uri = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize";
logout_uri = "https://login.microsoftonline.com/logout.srf";

url= "https://graph.microsoft.com/beta/groups?$expand=members";


windowWidth = 720;
windowHeight = 1024;

scope_prefix = "https://graph.microsoft.com/";
scopes = {
    "User.Read",
    "User.ReadBasic.All",
    "Tasks.Read",
    "Tasks.Read.Shared",
    "Group.Read.All"
};

[DataSource.Kind="MSPLANNERRESTAPI", Publish="MSPLANNERRESTAPI.UI"]
shared MSPLANNERRESTAPI.FunctionCatalog = () =>
let
    objects = #table(
        {"Name","Key","Data","ItemKind","ItemName","IsLeaf"},
        {
            /* Users */
            {"GetUser","GetUser",MSPLANNERRESTAPI.GetUser,"Function","GetUser",true},
            /* Groups */
            {"ListUsersDirectMemberships","ListUsersDirectMemberships",MSPLANNERRESTAPI.ListUsersDirectMemberships,"Function","ListUsersDirectMemberships",true},
            /* Plans */
            {"ListPlans","ListPlans",MSPLANNERRESTAPI.ListPlans,"Function","ListPlans",true},
            {"GetPlan","GetPlan",MSPLANNERRESTAPI.GetPlan,"Function","GetPlan",true},
            {"ListPlanTasks","ListPlanTasks",MSPLANNERRESTAPI.ListPlanTasks,"Function","ListPlanTasks",true},
            {"ListBuckets","ListBuckets",MSPLANNERRESTAPI.ListBuckets,"Function","ListBuckets",true},    
            /* Tasks */
            {"GetTask","GetTask",MSPLANNERRESTAPI.GetTask,"Function","GetTask",true},  
            {"GetTaskDetails","GetTaskDetails",MSPLANNERRESTAPI.GetTaskDetails,"Function","GetTaskDetails",true},  
            /* Planner Details*/
            {"GetPlanDetails","GetPlanDetails",MSPLANNERRESTAPI.GetPlanDetails,"Function","GetPlanDetails",true},
            /* Bucket */
            {"GetBucket","GetBucket",MSPLANNERRESTAPI.GetBucket,"Function","GetBucket",true},
            {"ListBucketTasks","ListBucketTasks",MSPLANNERRESTAPI.ListBucketTasks,"Function","ListBucketTasks",true},
            /* Assigned to task board format */            
            {"GetAssignedToTaskBoardFormat","GetAssignedToTaskBoardFormat",MSPLANNERRESTAPI.GetAssignedToTaskBoardFormat,"Function","GetAssignedToTaskBoardFormat",true},
            /* Bucket task board task format */
            {"GetBucketTaskBoardFormat","GetBucketTaskBoardFormat",MSPLANNERRESTAPI.GetBucketTaskBoardFormat,"Function","GetBucketTaskBoardFormat",true},
            /* Progress task board task format */
            {"GetProgressTaskBoardFormat","GetProgressTaskBoardFormat",MSPLANNERRESTAPI.GetProgressTaskBoardFormat,"Function","GetProgressTaskBoardFormat",true}            
        }),
    navTable = Table.ForceToNavigationTable(objects,{"Key"},"Name","Data","ItemKind","ItemName","IsLeaf")
in
    navTable;

/*****************
    Users
*****************/

/*** GetUser ***/
shared MSPLANNERRESTAPI.GetUser = Value.ReplaceType(
    GetUser,
    GetUserType
);

GetUserType = type function ( 
    optional groupId as ( 
        type text meta [
            Documentation.FieldCaption = "User ID",
            Documentation.FieldDescription = "The object ID or User Principal Name for the user",
            Documentation.SampleValues = {
                "288a6513-8516-4dc9-a790-3706510exxxx"
            }
        ]
    )
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.GetUser",
    Documentation.LongDescription = "Retrieve the properties and relationships of user object.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.GetUser(""288a6513-8516-4dc9-a790-3706510exxxx"")",
            Result = "If successful, this method a user object in the response body. https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

GetUser = (optional userId as text) =>
    let 
        // Check arguments
        doArgumentsExist = if(userId <> null and userId <> "") then true else false,
        // Build results
        result = if (doArgumentsExist = false)
                 then 
                    // Send back message that arguments are required
                    #table( type table [response = text],{{"Please provide userId argument"}})
                 else
                    let
                        jsonResult = GetJSON("/users/" & userId)
                    in 
                        jsonResult                   
    in
        result;             
/*** End GetUser ***/

/*****************
    Groups
*****************/

/*** ListUsersDirectMemberships ***/
shared MSPLANNERRESTAPI.ListUsersDirectMemberships = Value.ReplaceType(
    ListUsersDirectMemberships,
    ListUsersDirectMembershipsType
);

ListUsersDirectMembershipsType = type function ( 
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.ListUsersDirectMemberships",
    Documentation.LongDescription = "Get groups, directory roles, and administrative units that the user is a direct member of. This operation is not transitive.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.ListUsersDirectMemberships()",
            Result = "If successful, this method returns a collection of directoryObject objects in the response body. https://learn.microsoft.com/en-us/graph/api/user-list-memberof?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

ListUsersDirectMemberships = () =>
    let 
        result = GetJSON("/me/memberOf")             
    in
        result;             
/*** End ListPlan ***/

/*****************
    Plans
*****************/

/*** ListPlans ***/
shared MSPLANNERRESTAPI.ListPlans = Value.ReplaceType(
    ListPlans,
    ListPlansType
);

ListPlansType = type function ( 
    optional groupId as ( 
        type text meta [
            Documentation.FieldCaption = "Group ID",
            Documentation.FieldDescription = "The object ID for the Group ID",
            Documentation.SampleValues = {
                "798a9745-0295-4ecc-bbc5-d7756b195xxx"
            }
        ]
    )
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.ListPlans",
    Documentation.LongDescription = "Retrieve a list of plannerPlan objects owned by a group object.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.ListPlan(""798a9745-0295-4ecc-bbc5-d7756b195xxx"")",
            Result = "If successful, this method returns a collection of plannerPlan objects in the response body. https://learn.microsoft.com/en-us/graph/api/plannergroup-list-plans?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

ListPlans = (optional groupId as text) =>
    let 
        // Check arguments
        doArgumentsExist = if(groupId <> null and groupId <> "") then true else false,
        // Build results
        result = if (doArgumentsExist = false)
                 then 
                    // Send back message that arguments are required
                    #table( type table [response = text],{{"Please provide groupId argument"}})
                 else
                    let
                        jsonResult = GetJSON("/groups/" & groupId & "/planner/plans")
                    in 
                        jsonResult                   
    in
        result;             
/*** End ListPlan ***/

/*** GetPlan ***/
shared MSPLANNERRESTAPI.GetPlan = Value.ReplaceType(
    GetPlan,
    GetPlanType
);

GetPlanType = type function ( 
    optional planId as ( 
        type text meta [
            Documentation.FieldCaption = "Plan ID",
            Documentation.FieldDescription = "The Plan ID for Planner",
            Documentation.SampleValues = {
                "ezJEUXdkTEqQJML3awehdWUADxxx"
            }
        ]
    )
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.GetPlan",
    Documentation.LongDescription = "Retrieve the properties and relationships of a plannerplan object.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.GetPlan(""ezJEUXdkTEqQJML3awehdWUADxxx"")",
            Result = "If successful, this method returns a plannerPlan object in the response body. https://learn.microsoft.com/en-us/graph/api/plannerplan-get?view=graph-rest-1.0&tabs=http#response-1"
        ]
    }
];

GetPlan = (optional planId as text) =>
    let 
        // Check arguments
        doArgumentsExist = if(planId <> null and planId <> "") then true else false,
        // Build results
        result = if (doArgumentsExist = false)
                 then 
                    // Send back message that arguments are required
                    #table( type table [response = text],{{"Please provide planId argument"}})
                 else
                    let
                        jsonResult = GetJSON("/planner/plans/" & planId)
                    in 
                        jsonResult                   
    in
        result;    
/*** End GetPlan ***/

/*** End ListPlanTasks ***/
shared MSPLANNERRESTAPI.ListPlanTasks = Value.ReplaceType(
    ListPlanTasks,
    ListPlanTasksType
);

ListPlanTasksType = type function ( 
    optional planId as ( 
        type text meta [
            Documentation.FieldCaption = "Plan ID",
            Documentation.FieldDescription = "The Plan ID for Planner",
            Documentation.SampleValues = {
                "ezJEUXdkTEqQJML3awehdWUADxxx"
            }
        ]
    )
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.ListPlanTasks",
    Documentation.LongDescription = "Retrieve a list of plannerTask objects associated with a plannerPlan object.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.GetPlan(""ezJEUXdkTEqQJML3awehdWUADxxx"")",
            Result = "If successful, this method returns a 200 OK response code and a collection of plannerTask objects in the response body. https://learn.microsoft.com/en-us/graph/api/plannerplan-list-tasks?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

ListPlanTasks = (optional planId as text) =>
    let 
        // Check arguments
        doArgumentsExist = if(planId <> null and planId <> "") then true else false,
        // Build results
        result = if (doArgumentsExist = false)
                 then 
                    // Send back message that arguments are required
                    #table( type table [response = text],{{"Please provide planId argument"}})
                 else
                    let
                        jsonResult = GetJSON("/planner/plans/" & planId & "/tasks")
                    in 
                        jsonResult                   
    in
        result;    
/*** End ListPlanTasks ***/

/*** ListBuckets ***/
shared MSPLANNERRESTAPI.ListBuckets = Value.ReplaceType(
    ListBuckets,
    ListBucketsType
);

ListBucketsType = type function ( 
    optional planId as ( 
        type text meta [
            Documentation.FieldCaption = "Plan ID",
            Documentation.FieldDescription = "The Plan ID for Planner",
            Documentation.SampleValues = {
                "ezJEUXdkTEqQJML3awehdWUADxxx"
            }
        ]
    )
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.ListBuckets",
    Documentation.LongDescription = "Retrieve a list of plannerBucket objects contained by a plannerPlan object.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.ListBuckets(""ezJEUXdkTEqQJML3awehdWUADxxx"")",
            Result = "If successful, this method returns a collection of plannerBucket objects in the response body. https://learn.microsoft.com/en-us/graph/api/plannerplan-list-buckets?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

ListBuckets = (optional planId as text) =>
    let 
        // Check arguments
        doArgumentsExist = if(planId <> null and planId <> "") then true else false,
        // Build results
        result = if (doArgumentsExist = false)
                 then 
                    // Send back message that arguments are required
                    #table( type table [response = text],{{"Please provide planId argument"}})
                 else
                    let
                        jsonResult = GetJSON("/planner/plans/" & planId & "/buckets")
                    in 
                        jsonResult                   
    in
        result;             
/*** End ListBuckets ***/

/*****************
    Tasks
*****************/

/*** GetTask ***/
shared MSPLANNERRESTAPI.GetTask = Value.ReplaceType(
    GetTask,
    GetTaskType
);

GetTaskType = type function ( 
    optional taskId as ( 
        type text meta [
            Documentation.FieldCaption = "Task ID",
            Documentation.FieldDescription = "The unique identifier to the task in Planner",
            Documentation.SampleValues = {
                "VBvYjVZB4k-H_xJq8TO2NmUAAxxx"
            }
        ]
    )
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.GetTask",
    Documentation.LongDescription = "Retrieve the properties and relationships of plannerTask object.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.GetTask(""VBvYjVZB4k-H_xJq8TO2NmUAAxxx"")",
            Result = "If successful, this method returns a plannerTask object in the response body. https://learn.microsoft.com/en-us/graph/api/plannertask-get?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

GetTask = (optional taskId as text) =>
    let 
        // Check arguments
        doArgumentsExist = if(taskId <> null and taskId <> "") then true else false,
        // Build results
        result = if (doArgumentsExist = false)
                 then 
                    // Send back message that arguments are required
                    #table( type table [response = text],{{"Please provide taskId argument"}})
                 else
                    let
                        jsonResult = GetJSON("/planner/tasks/" & taskId)
                    in 
                        jsonResult                   
    in
        result;             
/*** End GetTask ***/

/*** GetTaskDetails ***/
shared MSPLANNERRESTAPI.GetTaskDetails = Value.ReplaceType(
    GetTaskDetails,
    GetTaskDetailsType
);

GetTaskDetailsType = type function ( 
    optional taskId as ( 
        type text meta [
            Documentation.FieldCaption = "Task ID",
            Documentation.FieldDescription = "The unique identifier to the task in Planner",
            Documentation.SampleValues = {
                "VBvYjVZB4k-H_xJq8TO2NmUAAxxx"
            }
        ]
    )
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.GetTaskDetails",
    Documentation.LongDescription = "Retrieve the properties and relationships of a plannerTaskDetails object.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.GetTaskDetails(""VBvYjVZB4k-H_xJq8TO2NmUAAxxx"")",
            Result = "If successful, this method returns a plannerTaskDetails object in the response body. https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-get?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

GetTaskDetails = (optional taskId as text) =>
    let 
        // Check arguments
        doArgumentsExist = if(taskId <> null and taskId <> "") then true else false,
        // Build results
        result = if (doArgumentsExist = false)
                 then 
                    // Send back message that arguments are required
                    #table( type table [response = text],{{"Please provide taskId argument"}})
                 else
                    let
                        jsonResult = GetJSON("/planner/tasks/" & taskId & "/details")
                    in 
                        jsonResult                   
    in
        result;             
/*** End GetTaskDetails ***/

/*****************
    Planner Details
*****************/

/*** GetPlanDetails ***/
shared MSPLANNERRESTAPI.GetPlanDetails = Value.ReplaceType(
    GetPlanDetails,
    GetPlanDetailsType
);

GetPlanDetailsType = type function ( 
    optional planId as ( 
        type text meta [
            Documentation.FieldCaption = "Plan ID",
            Documentation.FieldDescription = "The Plan ID for Planner",
            Documentation.SampleValues = {
                "ezJEUXdkTEqQJML3awehdWUADxxx"
            }
        ]
    )
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.GetPlanDetails",
    Documentation.LongDescription = "Retrieve the properties and relationships of a plannerPlanDetails object.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.GetPlanDetails(""ezJEUXdkTEqQJML3awehdWUADxxx"")",
            Result = "If successful, this method returns a plannerPlanDetails object in the response body. https://learn.microsoft.com/en-us/graph/api/plannerplandetails-get?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

GetPlanDetails = (optional planId as text) =>
    let 
        // Check arguments
        doArgumentsExist = if(planId <> null and planId <> "") then true else false,
        // Build results
        result = if (doArgumentsExist = false)
                 then 
                    // Send back message that arguments are required
                    #table( type table [response = text],{{"Please provide planId argument"}})
                 else
                    let
                        jsonResult = GetJSON("/planner/plans/" & planId & "/details")
                    in 
                        jsonResult                   
    in
        result;    
/*** End GetPlanDetails ***/

/*****************
    Buckets
*****************/

/*** GetBucket ***/
shared MSPLANNERRESTAPI.GetBucket = Value.ReplaceType(
    GetBucket,
    GetBucketType
);

GetBucketType = type function ( 
    optional bucketId as ( 
        type text meta [
            Documentation.FieldCaption = "Bucket ID",
            Documentation.FieldDescription = "The unique identifer for the Bucket for Planner",
            Documentation.SampleValues = {
                "ezJEUXdkTEqQJML3awehdWUADxxx"
            }
        ]
    )
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.GetBucket",
    Documentation.LongDescription = "Retrieve the properties and relationships of a plannerBucket object.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.GetBucket(""ezJEUXdkTEqQJML3awehdWUADxxx"")",
            Result = "If successful, this method returns a plannerBucket object in the response body. https://learn.microsoft.com/en-us/graph/api/plannerbucket-get?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

GetBucket = (optional bucketId as text) =>
    let 
        // Check arguments
        doArgumentsExist = if(bucketId <> null and bucketId <> "") then true else false,
        // Build results
        result = if (doArgumentsExist = false)
                 then 
                    // Send back message that arguments are required
                    #table( type table [response = text],{{"Please provide bucketId argument"}})
                 else
                    let
                        jsonResult = GetJSON("/planner/buckets/" & bucketId)
                    in 
                        jsonResult                   
    in
        result;    
/*** End GetBucket ***/

/*** ListBucketTasks ***/
shared MSPLANNERRESTAPI.ListBucketTasks = Value.ReplaceType(
    ListBucketTasks,
    ListBucketTasksType
);

ListBucketTasksType = type function ( 
    optional bucketId as ( 
        type text meta [
            Documentation.FieldCaption = "Bucket ID",
            Documentation.FieldDescription = "The unique identifer for the Bucket for Planner",
            Documentation.SampleValues = {
                "ezJEUXdkTEqQJML3awehdWUADxxx"
            }
        ]
    )
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.ListBucketTasks",
    Documentation.LongDescription = "Retrieve a list of plannerTask objects associated to a plannerBucket object.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.ListBucketTasks(""ezJEUXdkTEqQJML3awehdWUADxxx"")",
            Result = "If successful, this method returns a collection of plannerTask objects in the response body. https://learn.microsoft.com/en-us/graph/api/plannerbucket-list-tasks?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

ListBucketTasks = (optional bucketId as text) =>
    let 
        // Check arguments
        doArgumentsExist = if(bucketId <> null and bucketId <> "") then true else false,
        // Build results
        result = if (doArgumentsExist = false)
                 then 
                    // Send back message that arguments are required
                    #table( type table [response = text],{{"Please provide bucketId argument"}})
                 else
                    let
                        jsonResult = GetJSON("/planner/buckets/" & bucketId & "/tasks")
                    in 
                        jsonResult                   
    in
        result;    
/*** End ListBucketTasks ***/

/*********************************
    Assigned to task board format
**********************************/

/*** GetAssignedToTaskBoardFormat ***/
shared MSPLANNERRESTAPI.GetAssignedToTaskBoardFormat = Value.ReplaceType(
    GetAssignedToTaskBoardFormat,
    GetAssignedToTaskBoardFormatType
);

GetAssignedToTaskBoardFormatType = type function ( 
    optional taskId as ( 
        type text meta [
            Documentation.FieldCaption = "Task ID",
            Documentation.FieldDescription = "The unique identifier to the task in Planner",
            Documentation.SampleValues = {
                "VBvYjVZB4k-H_xJq8TO2NmUAAxxx"
            }
        ]
    )
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.GetAssignedToTaskBoardFormat",
    Documentation.LongDescription = "Retrieve the properties and relationships of a plannerAssignedToTaskBoardTaskFormat object.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.GetAssignedToTaskBoardFormat(""VBvYjVZB4k-H_xJq8TO2NmUAAxxx"")",
            Result = "If successful, this method returns a plannerAssignedToTaskBoardTaskFormat object in the response body. https://learn.microsoft.com/en-us/graph/api/plannerassignedtotaskboardtaskformat-get?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

GetAssignedToTaskBoardFormat = (optional taskId as text) =>
    let 
        // Check arguments
        doArgumentsExist = if(taskId <> null and taskId <> "") then true else false,
        // Build results
        result = if (doArgumentsExist = false)
                 then 
                    // Send back message that arguments are required
                    #table( type table [response = text],{{"Please provide taskId argument"}})
                 else
                    let
                        jsonResult = GetJSON("/planner/tasks/" & taskId & "/assignedToTaskBoardFormat")
                    in 
                        jsonResult                   
    in
        result;             
/*** End GetAssignedToTaskBoardFormat ***/

/*********************************
    Bucket task board task format
**********************************/

/*** GetBucketTaskBoardFormat ***/
shared MSPLANNERRESTAPI.GetBucketTaskBoardFormat = Value.ReplaceType(
    GetBucketTaskBoardFormat,
    GetBucketTaskBoardFormatType
);

GetBucketTaskBoardFormatType = type function ( 
    optional taskId as ( 
        type text meta [
            Documentation.FieldCaption = "Task ID",
            Documentation.FieldDescription = "The unique identifier to the task in Planner",
            Documentation.SampleValues = {
                "VBvYjVZB4k-H_xJq8TO2NmUAAxxx"
            }
        ]
    )
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.GetBucketTaskBoardFormat",
    Documentation.LongDescription = "Retrieve the properties and relationships of plannerBucketTaskBoardTaskFormat object.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.GetBucketTaskBoardFormat(""VBvYjVZB4k-H_xJq8TO2NmUAAxxx"")",
            Result = "If successful, this method returns a plannerBucketTaskBoardTaskFormat object in the response body. https://learn.microsoft.com/en-us/graph/api/plannerbuckettaskboardtaskformat-get?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

GetBucketTaskBoardFormat = (optional taskId as text) =>
    let 
        // Check arguments
        doArgumentsExist = if(taskId <> null and taskId <> "") then true else false,
        // Build results
        result = if (doArgumentsExist = false)
                 then 
                    // Send back message that arguments are required
                    #table( type table [response = text],{{"Please provide taskId argument"}})
                 else
                    let
                        jsonResult = GetJSON("/planner/tasks/" & taskId & "/bucketTaskBoardFormat")
                    in 
                        jsonResult                   
    in
        result;             
/*** End GetBucketTaskBoardFormat ***/

/*********************************
    Progress task board task format
**********************************/

/*** GetProgressTaskBoardFormat ***/
shared MSPLANNERRESTAPI.GetProgressTaskBoardFormat = Value.ReplaceType(
    GetProgressTaskBoardFormat,
    GetProgressTaskBoardFormatType
);

GetProgressTaskBoardFormatType = type function ( 
    optional taskId as ( 
        type text meta [
            Documentation.FieldCaption = "Task ID",
            Documentation.FieldDescription = "The unique identifier to the task in Planner",
            Documentation.SampleValues = {
                "VBvYjVZB4k-H_xJq8TO2NmUAAxxx"
            }
        ]
    )
) as table meta [
    Documentation.Name = "MSPLANNERRESTAPI.GetProgressTaskBoardFormat",
    Documentation.LongDescription = "Retrieve the properties and relationships of plannerProgressTaskBoardTaskFormat object.", 
    Documentation.Examples = {
        [
            Code = "=MSPLANNERRESTAPI.GetProgressTaskBoardFormat(""VBvYjVZB4k-H_xJq8TO2NmUAAxxx"")",
            Result = "If successful, this method returns a plannerProgressTaskBoardTaskFormat object in the response body. https://learn.microsoft.com/en-us/graph/api/plannerprogresstaskboardtaskformat-get?view=graph-rest-1.0&tabs=http#response"
        ]
    }
];

GetProgressTaskBoardFormat = (optional taskId as text) =>
    let 
        // Check arguments
        doArgumentsExist = if(taskId <> null and taskId <> "") then true else false,
        // Build results
        result = if (doArgumentsExist = false)
                 then 
                    // Send back message that arguments are required
                    #table( type table [response = text],{{"Please provide taskId argument"}})
                 else
                    let
                        jsonResult = GetJSON("/planner/tasks/" & taskId & "/progressTaskBoardFormat")
                    in 
                        jsonResult                   
    in
        result;             
/*** End GetProgressTaskBoardFormat ***/

/**** Help Functions ****/
GetJSON = (optional relPath as text) =>
    let
        params = [Headers = [#"Content-Type" = "application/json"], RelativePath = relPath],
        //Get Pages
        //Hard-coded Base URL because Power BI Service won't recognize parameter despite numerous attempts
        webSource = Web.Contents(
            "https://graph.microsoft.com/v1.0",
            params
        ),
        //Convert to Json
        result = Text.FromBinary(webSource)
    in
        result;
/**** End Helper Functions ****/


MSPLANNERRESTAPI = [
    TestConnection = (dataSourcePath) => { "MSPLANNERRESTAPI.FunctionCatalog" },
    Authentication = [
        OAuth = [
            StartLogin=StartLogin,
            FinishLogin=FinishLogin,
            Refresh=Refresh,
            Logout=Logout
        ]
    ],
    Label = "Microsoft Planner REST API Custom Connector"
];

// Setup Microsoft Planner Connector
MSPLANNERRESTAPI.UI = [
    Beta = true,
    Category = "Other",
    ButtonText = {
        Extension.LoadString(
            "ButtonTitle"
        ),
        Extension.LoadString(
            "ButtonHelp"
        )
    },    
    SourceImage = MSPLANNERRESTAPI.Icons,
    SourceTypeImage = MSPLANNERRESTAPI.Icons
];

MSPLANNERRESTAPI.Icons = [
    Icon16 = { Extension.Contents("powerquery-connector-ms-planner-rest-api16.png"), Extension.Contents("powerquery-connector-ms-planner-rest-api20.png"), Extension.Contents("powerquery-connector-ms-planner-rest-api24.png"), Extension.Contents("powerquery-connector-ms-planner-rest-api32.png") },
    Icon32 = { Extension.Contents("powerquery-connector-ms-planner-rest-api32.png"), Extension.Contents("powerquery-connector-ms-planner-rest-api40.png"), Extension.Contents("powerquery-connector-ms-planner-rest-api48.png"), Extension.Contents("powerquery-connector-ms-planner-rest-api64.png") }
];


// Thanks to https://aycabas.com/2020/05/11/graph-api-custom-connector/ for the artcle I found to do
// OAuth with Graph API.
StartLogin = (resourceUrl, state, display) =>
    let
        authorizeUrl = authorize_uri & "?" & Uri.BuildQueryString([
            client_id = client_id,  
            redirect_uri = redirect_uri,
            state = state,
            scope = "offline_access " & GetScopeString(scopes, scope_prefix),
            response_type = "code",
            response_mode = "query",
            login = "login"
        ])
    in
        [
            LoginUri = authorizeUrl,
            CallbackUri = redirect_uri,
            WindowHeight = 720,
            WindowWidth = 1024,
            Context = null
        ];


FinishLogin = (context, callbackUri, state) =>
    let
        parts = Uri.Parts(callbackUri)[Query],
 
        result = if (Record.HasFields(parts, {"error", "error_description"})) then 
                    error Error.Record(parts[error], parts[error_description], parts)
                 else
                    TokenMethod("authorization_code", "code", parts[code])
    in
        result;

// Refresh
// Original Signature: https://learn.microsoft.com/en-us/power-query/handling-authentication#original-oauth-signatures
Refresh = (resourceUrl, refresh_token) => TokenMethod("refresh_token", "refresh_token", refresh_token);

// Logout
// Original Signature: https://learn.microsoft.com/en-us/power-query/handling-authentication#original-oauth-signatures
Logout = (token) => logout_uri;

// Token Method
// Refresh Token
TokenMethod = (grantType, tokenField, code) =>
    let
        queryString = [
            client_id = client_id,
            scope = "offline_access " & GetScopeString(scopes, scope_prefix),
            grant_type = grantType,
            redirect_uri = redirect_uri
        ],
        queryWithCode = Record.AddField(queryString, tokenField, code),

        tokenResponse = Web.Contents(token_uri, [
            Content = Text.ToBinary(Uri.BuildQueryString(queryWithCode)),
            Headers = [
                #"Content-type" = "application/x-www-form-urlencoded",
                #"Accept" = "application/json"
            ],
            ManualStatusHandling = {400} 
        ]),
        body = Json.Document(tokenResponse),
        result = if (Record.HasFields(body, {"error", "error_description"})) then 
                    error Error.Record(body[error], body[error_description], body)
                 else
                    body
    in
        result;

// Handle scope for OAuth
Value.IfNull = (a, b) => if a <> null then a else b;

GetScopeString = (scopes as list, optional scopePrefix as text) as text =>
    let
        prefix = Value.IfNull(scopePrefix, ""),
        addPrefix = List.Transform(scopes, each prefix & _),
        asText = Text.Combine(addPrefix, " ")
    in
        asText;

//To Navigation Table
Table.ToNavigationTable = (
    table as table,
    keyColumns as list,
    nameColumn as text,
    dataColumn as text,
    itemKindColumn as text,
    itemNameColumn as text,
    isLeafColumn as text
) as table =>
    let
        tableType = Value.Type(
            table
        ),
        newTableType = Type.AddTableKey(
            tableType,
            keyColumns,
            true
        ) meta [
            NavigationTable.NameColumn = nameColumn,
            NavigationTable.DataColumn = dataColumn,
            NavigationTable.ItemKindColumn = itemKindColumn,
            Preview.DelayColumn = itemNameColumn,
            NavigationTable.IsLeafColumn = isLeafColumn
        ],
        navigationTable = Value.ReplaceType(
            table,
            newTableType
        )
    in
        navigationTable;

// Force Navigation Table
Table.ForceToNavigationTable = (
    table as table,
    keyColumns as list,
    nameColumn as text,
    dataColumn as text,
    itemKindColumn as text,
    itemNameColumn as text,
    isLeafColumn as text
) as table =>
    let
        tableType = Value.Type(
            table
        ),
        newTableType = Type.AddTableKey(
            tableType,
            keyColumns,
            true
        ) meta [
            NavigationTable.NameColumn = nameColumn,
            NavigationTable.DataColumn = dataColumn,
            NavigationTable.ItemKindColumn = itemKindColumn,
            NavigationTable.IsLeafColumn = isLeafColumn
        ],
        navigationTable = Value.ReplaceType(
            table,
            newTableType
        )
    in
        navigationTable;