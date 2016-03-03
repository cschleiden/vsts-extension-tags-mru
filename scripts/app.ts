/// <reference path='../typings/tsd.d.ts' />

import Q = require("q");

import TFS_Wit_Contracts = require("TFS/WorkItemTracking/Contracts"); 
import TFS_Wit_Client = require("TFS/WorkItemTracking/RestClient"); 
import TFS_Wit_Services = require("TFS/WorkItemTracking/Services");

import VSS_Extension_Service = require("VSS/SDK/Services/ExtensionData");

/** Maximum size of MRU */
const MAX_TAGS = 5;

class Tags {
    /** Key used for document service */
    private static KEY: string = "tags";

    private static instance: Tags = null;

    /** Get or create singleton instance */
    public static getInstance(): IPromise<Tags> {
        if (Tags.instance) {
            return Q(Tags.instance);
        } else {
            return VSS.getService(VSS.ServiceIds.ExtensionData).then((dataService: VSS_Extension_Service.ExtensionDataService) => {
                return dataService.getValue(Tags.KEY, {
                    defaultValue: [],
                    scopeType: "User"
                }).then((savedTags: string[]) => {
                    Tags.instance = new Tags(MAX_TAGS);

                    if (savedTags) {
                        savedTags.forEach(t => Tags.instance.addTag(t));
                    }

                    return Tags.instance;
                });
            });
        }
    }

    private queue: string[] = [];

    constructor(private maxCount: number) {
    }

    public addTag(tag: string) {
        // Remove tag from current position
        var idx = this.queue.indexOf(tag);
        if (idx !== -1) {
            this.queue.splice(idx, 1);
        }        

        // Add tag in first position
        this.queue.unshift(tag);

        this.prune();
    }

    public getTags(): string[] {
        return this.queue;
    }

    public persist(): IPromise<any> {
        return VSS.getService(VSS.ServiceIds.ExtensionData).then((dataService: VSS_Extension_Service.ExtensionDataService) => {
            dataService.setValue(Tags.KEY, this.queue, {
                scopeType: "User"                
            });
        });
    }

    /** Ensure only maximum number of tags configured is stored */
    private prune() {
        if (this.queue.length > this.maxCount) {
            for (var i = 0; i < this.queue.length - MAX_TAGS; ++i) {
                this.queue.pop();
            }
        }
    }
}

// Proactively initialize instance and load tags
Tags.getInstance();

// Register context menu action
VSS.register("tags-mru-work-item-menu", {
    getMenuItems: (context) => {
        // Not all areas use the same format for passing work item ids. "ids" for Queries
        // "workItemIds" for backlogs
        var ids = context.ids || context.workItemIds;
        if (!ids && context.id) {
            // Boards only support a single work item
            ids = [context.id];
        }

        let calledWithActiveForm = false;

        if (!ids && context.workItemId) {
            // Work item form menu
            ids = [context.workItemId];
            calledWithActiveForm = true;
        }

        return Tags.getInstance().then(tags => {
            var childItems: IContributedMenuItem[] = [];

            tags.getTags().forEach(tag => {
                childItems.push(<IContributedMenuItem>{
                    text: tag,
                    title: `Add tag: ${tag}`,
                    action: () => {
                        if (calledWithActiveForm) {                            
                            // Modify active work item
                            TFS_Wit_Services.WorkItemFormService.getService().then(wi => {
                                (<IPromise<string>>wi.getFieldValue("System.Tags")).then(changedTagsRaw => {
                                    let tags = splitTags(changedTagsRaw)
                                        .map(t => t.trim())
                                        .filter(t => !!t);
                                    
                                    if (tags.indexOf(tag) === -1) {                                    
                                        wi.setFieldValue("System.Tags", 
                                            tags.concat([tag])
                                                .join(";"));
                                    }
                                });
                            });
                        } else {
                            // Get work items, add the new tag to the list of existing tags, and then update
                            var client = TFS_Wit_Client.getClient();

                            client.getWorkItems(ids).then((workItems) => {
                                for (var workItem of workItems) {
                                    var prom = client.updateWorkItem([{
                                        "op": "add",
                                        "path": "/fields/System.Tags",
                                        "value": (workItem.fields["System.Tags"] || "") + ";" + tag
                                    }], workItem.id);
                                }
                            });
                        }
                    }
                });
            });

            if (childItems.length === 0) {
                childItems.push(<IContributedMenuItem>{
                    title: "No tag added",
                    disabled: true
                });
            }

            return [<IContributedMenuItem>{
                title: "Recent Tags",
                childItems: childItems
            }]
        });
    }
});

/**
 *  
 * Tags are stored as a single field, separated by ";". We need to keep track of the tags when a work item
 * was opened, and the ones when it's closed. The intersection are the tags added.
 */
class WorkItemTagsListener {
    private static instance: WorkItemTagsListener = null;
    
    public static getInstance(): WorkItemTagsListener {
        if (!WorkItemTagsListener.instance) {
            WorkItemTagsListener.instance = new WorkItemTagsListener();
        }
        
        return WorkItemTagsListener.instance;
    }
    
    /** Holds tags when work item was opened */
    private orgTags: { [id: number]: string[] } = {};
    
    /** Tags added  */
    private newTags: { [id: number]: string[] } = {};
    
    public setOriginalTags(workItemId: number, tags: string[]) {
        this.orgTags[workItemId] = tags;
    }
    
    public setNewTags(workItemId: number, tags: string[]) {
        this.newTags[workItemId] = tags;
    }    
    
    public clearForWorkItem(workItemId: number) {
        delete this.orgTags[workItemId];
        delete this.newTags[workItemId];
    }
    
    public commitTagsForWorkItem(workItemId: number): IPromise<any> {
        return Tags.getInstance().then(tags => {
            // Generate intersection between old and new tags
            var orgTags = this.orgTags[workItemId] || [];
            
            var diffTags = (this.newTags[workItemId] || []).filter(t => orgTags.indexOf(t) < 0);

            for (var tag of diffTags) {
                if (!tag) {
                    continue;
                }

                tags.addTag(tag);
            }

            // Save tags to server
            tags.persist();
        });
    }    
}

function splitTags(rawTags: string): string[] {
    return rawTags.split(";").map(t => t.trim());
}

// Register work item change listener
VSS.register("tags-mru-work-item-form-observer", (context) => {
    var setOriginalTags = (args) => {
        // Get original tags from work item
        TFS_Wit_Services.WorkItemFormService.getService().then(wi => {
            (<IPromise<string>>wi.getFieldValue("System.Tags")).then(changedTagsRaw => {                    
                WorkItemTagsListener.getInstance().setOriginalTags(args.id, splitTags(changedTagsRaw));
            });
        });
    };
    
    return {
        onFieldChanged: (args) => {
            if (args.changedFields["System.Tags"]) {
                var changedTagsRaw: string = args.changedFields["System.Tags"];
                WorkItemTagsListener.getInstance().setNewTags(args.id, splitTags(changedTagsRaw));
            }   
        },
        onLoaded: (args) => setOriginalTags,
        onUnloaded: (args) => {
            // When the users choses "Save & Close", unloaded is sometimes fired before the save event, so
            // do not clean for now.
            //WorkItemTagsListener.getInstance().clearForWorkItem(args.id);
        },
        onSaved: (args) => {
            WorkItemTagsListener.getInstance().commitTagsForWorkItem(args.id);
        },
        onReset: (args) => setOriginalTags,
        onRefreshed: (args) => setOriginalTags
    };
});
