define(["require", "exports", "q", "TFS/WorkItemTracking/RestClient", "TFS/WorkItemTracking/Services"], function (require, exports, Q, TFS_Wit_Client, TFS_Wit_Services) {
    var Tags = (function () {
        function Tags(maxCount) {
            this.maxCount = maxCount;
            this.dict = {};
            this.queue = [];
        }
        Tags.getInstance = function () {
            if (Tags.instance) {
                return Q(Tags.instance);
            }
            else {
                return VSS.getService(VSS.ServiceIds.ExtensionData).then(function (dataService) {
                    return dataService.getValue(Tags.KEY, {
                        defaultValue: [],
                        scopeType: "User"
                    }).then(function (savedTags) {
                        Tags.instance = new Tags(MAX_TAGS);
                        if (savedTags) {
                            savedTags.forEach(function (t) { return Tags.instance.addTag(t); });
                        }
                        return Tags.instance;
                    });
                });
            }
        };
        Tags.prototype.addTag = function (tag) {
            if (this.dict[tag]) {
                var idx = this.queue.indexOf(tag);
                this.queue.splice(idx, 1);
            }
            else {
                this.dict[tag] = true;
            }
            this.queue.unshift(tag);
            this.prune();
        };
        Tags.prototype.getTags = function () {
            return this.queue;
        };
        Tags.prototype.persist = function () {
            var _this = this;
            return VSS.getService(VSS.ServiceIds.ExtensionData).then(function (dataService) {
                dataService.setValue(Tags.KEY, _this.queue, {
                    scopeType: "User"
                });
            });
        };
        Tags.prototype.prune = function () {
            if (this.queue.length > this.maxCount) {
                for (var i = 0; i < this.queue.length - MAX_TAGS; ++i) {
                    delete this.dict[this.queue.pop()];
                }
            }
        };
        Tags.KEY = "tags";
        Tags.instance = null;
        return Tags;
    })();
    var MAX_TAGS = 5;
    Tags.getInstance();
    VSS.register("tags-mru-work-item-menu", {
        getMenuItems: function (context) {
            var ids = context.ids || context.workItemIds;
            if (!ids && context.id) {
                ids = [context.id];
            }
            return Tags.getInstance().then(function (tags) {
                var childItems = [];
                tags.getTags().forEach(function (tag) {
                    childItems.push({
                        text: tag,
                        title: "Add tag: " + tag,
                        action: function () {
                            var client = TFS_Wit_Client.getClient();
                            client.getWorkItems(ids).then(function (workItems) {
                                for (var _i = 0; _i < workItems.length; _i++) {
                                    var workItem = workItems[_i];
                                    var prom = client.updateWorkItem([{
                                            "op": "add",
                                            "path": "/fields/System.Tags",
                                            "value": (workItem.fields["System.Tags"] || "") + ";" + tag
                                        }], workItem.id);
                                }
                            });
                        }
                    });
                });
                if (childItems.length === 0) {
                    childItems.push({
                        title: "No tag added",
                        disabled: true
                    });
                }
                return [{
                        title: "Recent Tags",
                        childItems: childItems
                    }];
            });
        }
    });
    var WorkItemTagsListener = (function () {
        function WorkItemTagsListener() {
            this.orgTags = {};
            this.newTags = {};
        }
        WorkItemTagsListener.getInstance = function () {
            if (!WorkItemTagsListener.instance) {
                WorkItemTagsListener.instance = new WorkItemTagsListener();
            }
            return WorkItemTagsListener.instance;
        };
        WorkItemTagsListener.prototype.setOriginalTags = function (workItemId, tags) {
            this.orgTags[workItemId] = tags;
        };
        WorkItemTagsListener.prototype.setNewTags = function (workItemId, tags) {
            this.newTags[workItemId] = tags;
        };
        WorkItemTagsListener.prototype.clearForWorkItem = function (workItemId) {
            delete this.orgTags[workItemId];
            delete this.newTags[workItemId];
        };
        WorkItemTagsListener.prototype.commitTagsForWorkItem = function (workItemId) {
            var _this = this;
            return Tags.getInstance().then(function (tags) {
                var diffTags = _this.newTags[workItemId].filter(function (t) { return _this.orgTags[workItemId].indexOf(t) < 0; });
                for (var _i = 0; _i < diffTags.length; _i++) {
                    var tag = diffTags[_i];
                    if (!tag) {
                        continue;
                    }
                    tags.addTag(tag);
                }
                _this.clearForWorkItem(workItemId);
                tags.persist();
            });
        };
        WorkItemTagsListener.instance = null;
        return WorkItemTagsListener;
    })();
    function splitTags(rawTags) {
        return rawTags.split(";").map(function (t) { return t.trim(); });
    }
    VSS.register("tags-mru-work-item-form-observer", function (context) {
        return {
            onFieldChanged: function (args) {
                if (args.changedFields["System.Tags"]) {
                    var changedTagsRaw = args.changedFields["System.Tags"];
                    WorkItemTagsListener.getInstance().setNewTags(args.id, splitTags(changedTagsRaw));
                }
            },
            onLoaded: function (args) {
                TFS_Wit_Services.WorkItemFormService.getService().then(function (wi) {
                    wi.getFieldValue("System.Tags").then(function (changedTagsRaw) {
                        WorkItemTagsListener.getInstance().setOriginalTags(args.id, splitTags(changedTagsRaw));
                    });
                });
            },
            onUnloaded: function (args) {
                WorkItemTagsListener.getInstance().clearForWorkItem(args.id);
            },
            onSaved: function (args) {
                WorkItemTagsListener.getInstance().commitTagsForWorkItem(args.id);
            },
            onReset: function (args) {
                WorkItemTagsListener.getInstance().clearForWorkItem(args.id);
            },
            onRefreshed: function (args) {
                WorkItemTagsListener.getInstance().clearForWorkItem(args.id);
            }
        };
    });
});
