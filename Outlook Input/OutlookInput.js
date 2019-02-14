
var cnfg = {
    RecordLimit: 1,
    UserName: 'rpaugh82@live.com',
    Password: 'Dis@ll0w3d',
    ExchangeServerVersion: getExchangeVersion(7),
    UseManualServiceURL: false,
    ServiceURL: null,
    UseDifferentMailbox: false,
    Mailbox: null,
    Folder: 4/*getWellKnownFolderName(4)*/,
    IncludeRecurringEvents: true,
    StartDate: '2018-01-01',
    EndDate: '2018-12-31',
    AttachmentPath: null,
    QueryString: 'subject:test',
    IncludeSubFolders: true,
    SubFolderName: null,
    IncludeFolderNameInOutput: true,
    SkipRootFolder: false,
    //UseUniqueFileName: false,
    AttachmentFilter: 'csv',
    Fields: [
        //EwsJS.ItemSchema.Id,
        EwsJS.EmailMessageSchema["Id"],
        EwsJS.EmailMessageSchema["Subject"],
        EwsJS.EmailMessageSchema["DateTimeReceived"],
        EwsJS.EmailMessageSchema["HasAttachments"],
        EwsJS.EmailMessageSchema["Attachments"],
        /*EwsJS.AppointmentSchema.Start,
        EwsJS.AppointmentSchema.End*/
        EwsJS.EmailMessageSchema["From"]
    ]
}

//var c;


//Alteryx.Engine.SendMessage.Info("c.Configuration");
//Alteryx.Engine.SendMessage.Info(c.Configuration);

function GetItems() {
    Alteryx.Engine.SendMessage.Info("c:");
    Alteryx.Engine.SendMessage.Info(c);
    //Alteryx.Engine.SendMessage.Info(c);
    //OutlookInput = c;
    //Alteryx.Engine.SendMessage.Info(OutlookInput);

    /*Promise.all([
        someProcedure(5).then(x => Alteryx.Engine.SendMessage.Info(x)),
        someProcedure(10).then(x => Alteryx.Engine.SendMessage.Info(x))
    ])
    .then(([result1, result2]) => {
        Alteryx.Engine.SendMessage.Info(result1);
        Alteryx.Engine.SendMessage.Info(result2);
        Alteryx.Engine.SendMessage.Complete();
    })
    .catch(err => {
        // Receives first rejection among the Promises
        Alteryx.Engine.SendMessage.Info("Promise All Error:");
        Alteryx.Engine.SendMessage.Info(err);
    });*/

    Alteryx.Engine.SendMessage.Info("Inbox Folder:");
    Alteryx.Engine.SendMessage.Info(c.FolderToSearch);
    Alteryx.Engine.SendMessage.Info(mailbox);
    Alteryx.Engine.SendMessage.Info(new EwsJS.FolderId(Number(c.FolderToSearch), mailbox));
    Alteryx.Engine.SendMessage.Info("GetItems() method reached.");
    Alteryx.Engine.SendMessage.Info("Use Manual Service URL? " + c.UseManualServiceURL);
    //set ews endpoint url to use
    // TODO: incorporate the http to https redirect in the autodiscover service implementation.
    if (c.UseManualServiceURL == "True") {
        service.Url = new EwsJS.Uri(c.ServiceURL);
    } else {
        Alteryx.Engine.SendMessage.Info("Before service url set:");
        // Switch to autodiscover.
        //service.AutodiscoverUrl("rpaugh82@live.com");
        //console.log(service.Url);
        service.Url = new EwsJS.Uri("https://outlook.office365.com/Ews/Exchange.asmx"); // you can also use exch.AutodiscoverUrl
        Alteryx.Engine.SendMessage.Info("after service url set:");
    }

    Alteryx.Engine.SendMessage.Info("Before set service:");
    //OutlookInput.Service = service;
    Alteryx.Engine.SendMessage.Info("After set service:");
    //Alteryx.Engine.SendMessage.Info(OutlookInput.Service.Url);

    var folderView = new EwsJS.FolderView(1000);
    folderView.PropertySet = new EwsJS.PropertySet(EwsJS.BasePropertySet.IdOnly);
    folderView.Traversal = EwsJS.FolderTraversal.Deep;
    
    //Promise.all([
        //!c.SkipRootFolder ? findItems([new EwsJS.FolderId(c.FolderToSearch, mailbox)]).then(x => x) : null,
        //findItems([new EwsJS.FolderId(c.FolderToSearch, mailbox)]).then(x => Alteryx.Engine.SendMessage.Info(x)),
        //someProcedure(5).then(x => Alteryx.Engine.SendMessage.Info(x)),
        FindFolders(folderView)
        .then(function (folders) {
            Alteryx.Engine.SendMessage.Info("Finding Folders:");
            Alteryx.Engine.SendMessage.Info(folders.length);
            //Alteryx.Engine.SendMessage.Info(folders);
            //Alteryx.Engine.SendMessage.Complete();
            
            return findItems(folders).then(x => Alteryx.Engine.SendMessage.Info(x));
        })
        .then(function (itemResults) {
            //Alteryx.Engine.SendMessage.Info("Item Results:");
            //Alteryx.Engine.SendMessage.Info(itemResults);
            Alteryx.Engine.SendMessage.Complete();
            //return itemResults;
        })
        .catch(function (error) {
            Alteryx.Engine.SendMessage.Error("FindFolders Parent Error");
            Alteryx.Engine.SendMessage.Error(error);
            Alteryx.Engine.SendMessage.Complete();
        });
        //someProcedure(10).then(x => Alteryx.Engine.SendMessage.Info(x))
    /*])
    .then(([result1, result2]) => {
        //Alteryx.Engine.SendMessage.Info("Promise All - Result 1:");
        //Alteryx.Engine.SendMessage.Info(result1);
        //Alteryx.Engine.SendMessage.Info("Promise All - Result 2:");
        //Alteryx.Engine.SendMessage.Info(result2);
        Alteryx.Engine.SendMessage.Complete();
    })
    .catch(err => {
        // Receives first rejection among the Promises
        Alteryx.Engine.SendMessage.Info("Promise All Error:");
        Alteryx.Engine.SendMessage.Info(err);
        Alteryx.Engine.SendMessage.Complete();
    });*/

    // Working
    /*FindFolders(folderView)
    .then(function (folders) {
        Alteryx.Engine.SendMessage.Info("Finding Folders:");
        Alteryx.Engine.SendMessage.Info(folders.length);
        //Alteryx.Engine.SendMessage.Info(folders);
        //Alteryx.Engine.SendMessage.Complete();
        
        return findItems(folders).then(x => Alteryx.Engine.SendMessage.Info(x));
    })
    .then(function (itemResults) {
        Alteryx.Engine.SendMessage.Info("Item Results:");
        Alteryx.Engine.SendMessage.Info(itemResults);
        Alteryx.Engine.SendMessage.Complete();
    })
    .catch(function (error) {
        Alteryx.Engine.SendMessage.Error("FindFolders Parent Error");
        Alteryx.Engine.SendMessage.Error(error);
        Alteryx.Engine.SendMessage.Complete();
    });*/

    /*FindItems(new EwsJS.FolderId(c.FolderToSearch, mailbox), c.QueryString, [])
    .then(function (result) {
        Alteryx.Engine.SendMessage.Info("FindItems Promise Method.");
        Alteryx.Engine.SendMessage.Info(result);
        GetItemsFromFolder(service, new EwsJS.FolderId(c.FolderToSearch, mailbox), result)
        .then(function (response) {
            Alteryx.Engine.SendMessage.Info("Records Response: ");
            // Push to Message stream:
            Alteryx.Engine.SendMessage.Info(response.records);
            Alteryx.Engine.SendMessage.PushRecords("Message", response.records);

            return service.GetAttachments(response.attachmentIDs, null, null)
            .then(function (attachmentResponse) {
                //Alteryx.Engine.SendMessage.Info(attachmentResponse);
                return attachmentResponse;
            })
            .catch(function (error) {
                Alteryx.Engine.SendMessage.Error(error);
                Alteryx.Engine.SendMessage.Complete();
            });
        })
        .then(function (attachmentResponse) {
            Alteryx.Engine.SendMessage.Info("Attachment Response: ");
            // Push to Attachment stream:
            Alteryx.Engine.SendMessage.Info(attachmentResponse.responses);
            var attachmentRecords = [];
            if (attachmentResponse.responses != undefined && attachmentResponse.responses.length > 0) {
                attachmentResponse.responses.forEach(function (attachment) {
                    Alteryx.Engine.SendMessage.Info("Attachment:");
                    Alteryx.Engine.SendMessage.Info(attachment.attachment.name);
                    Alteryx.Engine.SendMessage.Info(attachment.attachment.Base64Content);
                    attachmentRecords.push([
                        attachment.attachment.name,
                        attachment.attachment.Base64Content
                    ]);
                    Alteryx.Engine.SendMessage.Info("After array push");
                });
                Alteryx.Engine.SendMessage.Info(attachmentRecords);
                Alteryx.Engine.SendMessage.PushRecords("Attachments", attachmentRecords);
            }
            Alteryx.Engine.SendMessage.Complete();
        })
        .catch(function (error) {
            Alteryx.Engine.SendMessage.Error(error);
            Alteryx.Engine.SendMessage.Complete();
        });
    })
    .then(() => {
        //Alteryx.Engine.SendMessage.Complete();
    })
    .catch(function (error) {
        Alteryx.Engine.SendMessage.Error(error);
        Alteryx.Engine.SendMessage.Complete();
    });*/

    if (!c.SkipRootFolder) {
        // Get items from the selected root folder.
        // C#: GetItemsFromFolder(service, new FolderId(Folder mailbox), true);
            // Todo: replace hard-coded folder Id with user-provided value.
        //Alteryx.Engine.SendMessage.Info("Skipping root folder? " + c.SkipRootFolder);
        //GetItemsFromFolder(service, new EwsJS.FolderId(c.FolderToSearch, mailbox), c.QueryString, true);
        //Alteryx.Engine.SendMessage.Info("Post First GetItemsFromFolder() call");
    }

    /*var folderView = new EwsJS.FolderView(1000);
    folderView.PropertySet = new EwsJS.PropertySet(EwsJS.BasePropertySet.IdOnly);
    folderView.Traversal = EwsJS.FolderTraversal.Deep;

    //Alteryx.Engine.SendMessage.Info("Including sub folders? " + c.IncludeSubFolders);
    if (c.IncludeSubFolders) {
        //GetFolders(folderView);
    }*/
}

function GetFolders(folderView) {
    // Cycle through sub folders.

    //Alteryx.Engine.SendMessage.Info("GetFolders() Method Reached.");
    if (c.SubFolderName != "" && c.SubFolderName != null) {
        // Search specific sub-folder.
        // Todo: replace hard-coded folder Id with user-provided value.
        service.FindFolders(c.FolderToSearch, new EwsJS.SearchFilter.ContainsSubstring(EwsJS.FolderSchema.DisplayName, c.SubFolderName), folderView)
        .then(function (folderResults) {
            for (var folder of folderResults.Folders) {
                //console.log(folder.propertyBag.properties.objects.Id);
                if (c.IncludeFolderNameInOutput) {
                    EwsJS.Folder.Bind(service, folder.propertyBag.properties.objects.Id, new EwsJS.PropertySet(EwsJS.BasePropertySet.IdOnly, [EwsJS.FolderSchema.DisplayName]))
                    .then(function (boundFolder) {
                        //console.log(boundFolder.Id);
                        //console.log(boundFolder.DisplayName);
                        //console.log("Bound Item:");
                        //console.log(boundFolder.Id);
                        //console.log(boundFolder.DisplayName);
                        GetItemsFromFolder(service, boundFolder, c.QueryString, false);
                    });
                } else {
                    GetItemsFromFolder(service, folder.propertyBag.properties.objects.Id, c.QueryString, false);
                }
                //console.log(folderResults.Folders[0].propertyBag.properties.objects.Id);
            }

            if (folderResults.MoreAvailable === true) {
                folderView.Offset += Array.from(folderResults.Folders).length;
    
                GetFolders(folderView);
            }
        });
    } else {
        // Search all sub-folders.
        // Todo: replace hard-coded folder Id with user-provided value.
        service.FindFolders(c.FolderToSearch, folderView)
        .then(function (folderResults) {
            for (var folder of folderResults.Folders) {
                //console.log(folderResults.Folders[0].propertyBag.properties.objects.Id);
                //GetItemsFromFolder(service, folder.propertyBag.properties.objects.Id, c.QueryString, false);
                if (c.IncludeFolderNameInOutput) {
                    EwsJS.Folder.Bind(service, folder.propertyBag.properties.objects.Id, new EwsJS.PropertySet(EwsJS.BasePropertySet.IdOnly, [EwsJS.FolderSchema.DisplayName]))
                    .then(function (boundFolder) {
                        //console.log(boundFolder.Id);
                        //console.log(boundFolder.DisplayName);
                        //console.log("Bound Item:");
                        //console.log(boundFolder.Id);
                        //console.log(boundFolder.DisplayName);
                        GetItemsFromFolder(service, boundFolder, c.QueryString, false);
                    });
                } else {
                    GetItemsFromFolder(service, folder.propertyBag.properties.objects.Id, c.QueryString, false);
                }
            }

            if (folderResults.MoreAvailable === true) {
                folderView.Offset += Array.from(folderResults.Folders).length;
    
                GetFolders(folderView);
            }
        });
    }
}

const someProcedure = async n =>
{
    for (let i = 0; i < n; i++) {
        const t = Math.random() * 1000;
        const x = await new Promise(r => setTimeout(r, t, i));
        Alteryx.Engine.SendMessage.Info(x);
    }
    return 'done';
}

const bindFolders = async r =>
{
    var folders = [];

    /*if (!c.SkipRootFolder) {
        folders.push(new EwsJS.FolderId(c.FolderToSearch, mailbox));
    }*/

    Alteryx.Engine.SendMessage.Info("bindFolders Count:");
    Alteryx.Engine.SendMessage.Info(r.Folders.length);
    for (let i = 0; i < r.Folders.length; i++) {
        //const x = await new Promise(r => setTimeout(r, t, i));
        
        if (c.IncludeFolderNameInOutput == "True") {
            //Alteryx.Engine.SendMessage.Info("Folder Id from bindFolders: " );
            //Alteryx.Engine.SendMessage.Info(r.Folders[i].propertyBag.properties.objects.Id);
            const x = await EwsJS.Folder.Bind(service, r.Folders[i].folderName != undefined ? r.Folders[i] : r.Folders[i].propertyBag.properties.objects.Id, new EwsJS.PropertySet(EwsJS.BasePropertySet.IdOnly, [EwsJS.FolderSchema.DisplayName]))
            .then(function (boundFolder) {
                Alteryx.Engine.SendMessage.Info("Bound Folder:");
                //Alteryx.Engine.SendMessage.Info(boundFolder.Id);
                Alteryx.Engine.SendMessage.Info(boundFolder.DisplayName);
                //Alteryx.Engine.SendMessage.Complete();
                folders.push(boundFolder);
            });
        } else {
            folders.push({ Id: r.Folders[i].propertyBag.properties.objects.Id });
        }
    }
    
    return folders;
}

const findItems = async folders =>
{    
    Alteryx.Engine.SendMessage.Info("Find Items Folder Count:");
    Alteryx.Engine.SendMessage.Info(folders.length);
    for (let i = 0; i < folders.length; i++) {
        //Alteryx.Engine.SendMessage.Info("Inside Item Loop: " + i);
        //Alteryx.Engine.SendMessage.Info(folders[i]);
        //Alteryx.Engine.SendMessage.Info(folders[i].folderName);
        //Alteryx.Engine.SendMessage.Info(t);
        //const x = await new Promise(r => setTimeout(r, t, i));
        const x = await FindItems(/*new EwsJS.FolderId(c.FolderToSearch, mailbox)*/folders[i], c.QueryString, new EwsJS.ItemView(1000), [])
        .then(function (findItemsResult) {
            Alteryx.Engine.SendMessage.Info("findItemResult Count (" + (folders[i].DisplayName == undefined ? getWellKnownFolderName(folders[i].folderName) : folders[i].DisplayName).toString() + "):");
            Alteryx.Engine.SendMessage.Info(findItemsResult.length);
            //Alteryx.Engine.SendMessage.Info("FindItems Promise Method.");
            //Alteryx.Engine.SendMessage.Info(findItemsResult);
            //Alteryx.Engine.SendMessage.Info("Folder:");
            //Alteryx.Engine.SendMessage.Info(folders[i].Id);
            //Alteryx.Engine.SendMessage.Info(folders[i].folderName);
            //Alteryx.Engine.SendMessage.PushRecords("Message", [["12321", "Test Subject", "2019-01-01", false, "From: Me"]]);
            if (findItemsResult != undefined && findItemsResult.length > 0) {
                return GetItemsFromFolder(service, /*new EwsJS.FolderId(c.FolderToSearch, mailbox)*/folders[i], findItemsResult)
                .then(function (response) {
                    Alteryx.Engine.SendMessage.Info("Records Response: ");
                    Alteryx.Engine.SendMessage.Info(response.records);
                    Alteryx.Engine.SendMessage.PushRecords("Message", response.records);
        
                    Alteryx.Engine.SendMessage.Info("Attachment IDs: ");
                    Alteryx.Engine.SendMessage.Info(response.attachmentIDs);
                    if (response.attachmentIDs != undefined && response.attachmentIDs.length > 0) {
                        return service.GetAttachments(response.attachmentIDs, null, null)
                        .then(function (attachmentResponse) {
                            //Alteryx.Engine.SendMessage.Info(attachmentResponse);
                            return attachmentResponse;
                        })
                        .catch(function (error) {
                            Alteryx.Engine.SendMessage.Error("Attachment Error");
                            Alteryx.Engine.SendMessage.Error(error);
                            Alteryx.Engine.SendMessage.Complete();
                        });
                    }

                    return;
                })
                .then(function (attachmentResponse) {
                    if (attachmentResponse != undefined) {
                        //Alteryx.Engine.SendMessage.Info("Attachment Response: ");
                        //Alteryx.Engine.SendMessage.Info(attachmentResponse);
                        //Alteryx.Engine.SendMessage.Info(attachmentResponse.responses);
                        var attachmentRecords = [];
                        if (attachmentResponse.responses != undefined && attachmentResponse.responses.length > 0) {
                            attachmentResponse.responses.forEach(function (attachment) {
                                Alteryx.Engine.SendMessage.Info("Attachment:");
                                Alteryx.Engine.SendMessage.Info(attachment.attachment.name);
                                //Alteryx.Engine.SendMessage.Info(attachment.attachment.Base64Content);
                                attachmentRecords.push([
                                    attachment.attachment.name,
                                    attachment.attachment.Base64Content
                                ]);
                                //Alteryx.Engine.SendMessage.Info("After array push");
                            });
                            //Alteryx.Engine.SendMessage.Info(attachmentRecords);
                            Alteryx.Engine.SendMessage.PushRecords("Attachments", attachmentRecords);
                        }
                        //Alteryx.Engine.SendMessage.Complete();
                    }

                    return;
                })
                .then(() => {
                    //return "Complete";
                })
                .catch(function (error) {
                    Alteryx.Engine.SendMessage.Error(error);
                    Alteryx.Engine.SendMessage.Complete();
                });
            }
        })
        .then(() => {
            //Alteryx.Engine.SendMessage.Complete();
        })
        .catch(function (error) {
            Alteryx.Engine.SendMessage.Error(error);
            Alteryx.Engine.SendMessage.Complete();
        });
    }

    return "Items Returned";
}

function FindFolders(folderView) {
    Alteryx.Engine.SendMessage.Info("Include Sub-Folders?");
    Alteryx.Engine.SendMessage.Info(c.IncludeSubFolders);
    //Alteryx.Engine.SendMessage.Info(OutlookInput.IncludeSubFolders);
    if (c.IncludeSubFolders == "True") {
        if (c.SubFolderName != "" && c.SubFolderName != null) {
            //Alteryx.Engine.SendMessage.Info("Inside First FindFolders Call.");
            // Search specific sub-folder.
            // Todo: replace hard-coded folder Id with user-provided value.
            return service.FindFolders(c.FolderToSearch, new EwsJS.SearchFilter.ContainsSubstring(EwsJS.FolderSchema.DisplayName, c.SubFolderName), folderView)
            .then(function (folderResults) {
                //Alteryx.Engine.SendMessage.Info("Find Folders Sub-Folder Search Promise Reached.");
                //Alteryx.Engine.SendMessage.Info(folderResults.Folders);
                if (c.SkipRootFolder == "False") {
                    folderResults.Folders.push(new EwsJS.FolderId(c.FolderToSearch, mailbox));
                }
                return bindFolders(folderResults).then(x => /*Alteryx.Engine.SendMessage.Info(x)*/x);
            })
            .then(function (loopResponse) {
                //Alteryx.Engine.SendMessage.Info("Loop Response:");
                //Alteryx.Engine.SendMessage.Info(loopResponse);
                return loopResponse;
            })
            .catch(function (error) {
                Alteryx.Engine.SendMessage.Error("Error inside service.FindFolders method (sub-folder).");
                Alteryx.Engine.SendMessage.Error(error);
                Alteryx.Engine.SendMessage.Complete();
            });
        } else {
            return service.FindFolders(c.FolderToSearch, folderView)
            .then(function (folderResults) {
                //Alteryx.Engine.SendMessage.Info("Find Folders Non-Sub-Folder Search Promise Reached.");
                //Alteryx.Engine.SendMessage.Info(folderResults);
                if (!c.SkipRootFolder) {
                    folderResults.Folders.push(new EwsJS.FolderId(c.FolderToSearch, mailbox));
                }
                return bindFolders(folderResults).then(x => /*Alteryx.Engine.SendMessage.Info(x)*/x);
            })
            .then(function (loopResponse) {
                //Alteryx.Engine.SendMessage.Info("Loop Response:");
                //Alteryx.Engine.SendMessage.Info(loopResponse);
                return loopResponse;
            })
            .catch(function (error) {
                Alteryx.Engine.SendMessage.Error("Error inside service.FindFolders method (parent folder).");
                Alteryx.Engine.SendMessage.Error(error);
                Alteryx.Engine.SendMessage.Complete();
            });
        }
    } else {
        Alteryx.Engine.SendMessage.Info("Inbox");
        var folderResults = {
            Folders: [
                new EwsJS.FolderId(c.FolderToSearch, mailbox)
            ]
        };
        Alteryx.Engine.SendMessage.Info("Do not include sub-folders.  Folder Results:");
        //Alteryx.Engine.SendMessage.Info(folderResults);
        return bindFolders(folderResults).then(x => /*Alteryx.Engine.SendMessage.Info(x)*/x);
    }
}

function FindItems(folder, queryString, itemView, itemIDs) {
    return service.FindItems(folder.Id == undefined ? folder : folder.Id, queryString, itemView)
    .then(function (result) {
        //var itemIDs = [];
        //Alteryx.Engine.SendMessage.Info("Iterate Item IDs");
        for (var item of result.Items) {
            //Alteryx.Engine.SendMessage.Info("Item Id: " + item.Id.UniqueId);
            itemIDs.push(item.Id);
        }

        Alteryx.Engine.SendMessage.Info("More Available (" + (folder.DisplayName == undefined ? getWellKnownFolderName(folder.folderName) : folder.DisplayName).toString() + "): " + result.MoreAvailable);
        Alteryx.Engine.SendMessage.Info("Item View Offset: " + itemView.Offset.toString());
        if (result.MoreAvailable === true) {
            itemView.Offset += Array.from(result.Items).length;

            return FindItems(/*c.SkipRootFolder ? folder.Id : folder*/folder, queryString, itemView, itemIDs);
        }

        return itemIDs;
    });
}

function GetItemsFromFolder(service, folder, itemIDs) {
    //Alteryx.Engine.SendMessage.Info("Inside GetItemsFromFolder() Method.");

    /*service.FindItems(folder.Id == undefined ? folder : folder.Id, queryString, itemView)
    .then(function (result) {
        var itemIDs = [];
        Alteryx.Engine.SendMessage.Info("Iterate Item IDs");
        for (var item of result.Items) {
            itemIDs.push(item.Id);
        }*/

        return service.BindToItems(itemIDs, new EwsJS.PropertySet(EwsJS.BasePropertySet.IdOnly, c.Fields))
        .then(function (responseItems) {
            //Alteryx.Engine.SendMessage.Info("Handle item responses");
            var records = [];
            var attachmentIDs = [];
            responseItems.responses.forEach(function (responseItem) {
                var record = [];
                if (c.IncludeFolderNameInOutput) {
                    record.push(folder.DisplayName == undefined ? getWellKnownFolderName(folder.folderName) : folder.DisplayName);
                }

                for (var field of c.Fields) {
                    Alteryx.Engine.SendMessage.Info("BindToItems field iteration:");
                    Alteryx.Engine.SendMessage.Info(field);
                    if (field.name === 'Id') {
                        record.push(responseItem.item.propertyBag.properties.objects[field.name].UniqueId);
                    } else {
                        if (field.name != 'Attachments') {
                            record.push(responseItem.item.propertyBag.properties.objects[field.name].toString());
                        }
                    }

                    //Alteryx.Engine.SendMessage.Info("Attachment Info");
                    //Alteryx.Engine.SendMessage.Info(field.name);
                    //Alteryx.Engine.SendMessage.Info(responseItem.item.propertyBag.properties.objects["HasAttachments"]);
                    if (field.name === 'Attachments' && responseItem.item.propertyBag.properties.objects["HasAttachments"]) {
                        Alteryx.Engine.SendMessage.Info("Attachments Found");
                        //console.log(boundItem.Attachments);
                        for (var attachment of responseItem.item.Attachments.items) {
                            Alteryx.Engine.SendMessage.Info("Attachment Id: " + attachment.Id);
                            if ((c.AttachmentFilter == null || c.AttachmentFilter == "") || (c.AttachmentFilter != null && attachment.Name.indexOf(c.AttachmentFilter) !== -1)) {
                                attachmentIDs.push(attachment.Id);
                            }
                        }
                    }
                }
                Alteryx.Engine.SendMessage.Info(record);

                records.push(record);
            });

            return {
                records: records,
                attachmentIDs: attachmentIDs
            };
            //Alteryx.Engine.SendMessage.Complete();
        })
        .catch(function (error) {
            Alteryx.Engine.SendMessage.Error("GetItemsFromFolder Error: ");
            Alteryx.Engine.SendMessage.Error(error);
            Alteryx.Engine.SendMessage.Complete();
        });
        //Alteryx.Engine.SendMessage.Info(records);
        //return records;
        /*Promise.all(records).then(function (response) {
            Alteryx.Engine.SendMessage.Info("Promise Response:");
            Alteryx.Engine.SendMessage.Info(reponse);
        });*/
        //Alteryx.Engine.SendMessage.Complete();
    /*})
    .then(function (response) {
        Alteryx.Engine.SendMessage.Info("Records Response: ");
        // Push to Message stream:
        Alteryx.Engine.SendMessage.Info(response.records);
        Alteryx.Engine.SendMessage.PushRecords("Message", response.records);

         return service.GetAttachments(response.attachmentIDs, null, null)
        .then(function (attachmentResponse) {
            //Alteryx.Engine.SendMessage.Info(attachmentResponse);
            return attachmentResponse;
        })
        .catch(function (error) {
            Alteryx.Engine.SendMessage.Error(error);
            Alteryx.Engine.SendMessage.Complete();
        });
    })
    .then(function (attachmentResponse) {
        Alteryx.Engine.SendMessage.Info("Attachment Response: ");
        // Push to Attachment stream:
        Alteryx.Engine.SendMessage.Info(attachmentResponse.responses);
        var attachmentRecords = [];
        if (attachmentResponse.responses != undefined && attachmentResponse.resopnses.length > 0) {
            attachmentResponse.responses.forEach(function (attachment) {
                Alteryx.Engine.SendMessage.Info("Attachment:");
                Alteryx.Engine.SendMessage.Info(attachment.attachment.name);
                Alteryx.Engine.SendMessage.Info(attachment.attachment.Base64Content);
                attachmentRecords.push([
                    attachment.attachment.name,
                    attachment.attachment.Base64Content
                ]);
                Alteryx.Engine.SendMessage.Info("After array push");
            });
            Alteryx.Engine.SendMessage.Info(attachmentRecords);
            Alteryx.Engine.SendMessage.PushRecords("Attachments", attachmentRecords);
        }
        Alteryx.Engine.SendMessage.Complete();
    })
    .catch(function (error) {
        Alteryx.Engine.SendMessage.Error(error);
        Alteryx.Engine.SendMessage.Complete();
    });*/

    /*service.FindItems(folder.Id == undefined ? folder : folder.Id, queryString, itemView)
    .then(function (results) {
        Alteryx.Engine.SendMessage.Info("First Promise Result.");
        //Alteryx.Engine.SendMessage.Info(JSON.stringify(results));
        for (var item of results.Items) {
            Alteryx.Engine.SendMessage.Info("Item Id: " + item.Id.UniqueId);
            EwsJS.Item.Bind(service, item.Id, new EwsJS.PropertySet(EwsJS.BasePropertySet.IdOnly, c.Fields))
            .then(function (boundItem) {
                Alteryx.Engine.SendMessage.Info("Bound Item Id: " + boundItem.propertyBag.properties.objects[field.Name].UniqueId);
                //console.log(bindResponse);
                //item.Load(new EwsJS.PropertySet(EwsJS.BasePropertySet.IdOnly, c.Fields)).then(() => {
                //console.log('Item Id: ' + item.Id.UniqueId);
                //let message = boundItem;
                var record = [];
                //record.push(c.SkipRootFolder ? f.DisplayName : getWellKnownFolderName(c.FolderToSearch));
                if (c.IncludeFolderNameInOutput) {
                    record.push(folder.DisplayName == undefined ? getWellKnownFolderName(folder.folderName) : folder.DisplayName);
                }

                for (var field of c.Fields) {
                    // Replace with Alteryx Engine message record output.
                    if (field.Name === 'Id') {
                        record.push(boundItem.propertyBag.properties.objects[field.Name].UniqueId);
                    } else {
                        if (field.Name != 'Attachments') {
                            record.push(boundItem.propertyBag.properties.objects[field.Name].toString());
                        }
                    }

                    // Replace with Alteryx Engine attachment record output.
                    if (field.Name === 'Attachments' && boundItem.HasAttachments) {
                        //console.log(boundItem.Attachments);
                        for (var attachment of boundItem.Attachments.items) {
                            //let file = boundItem.Attachments.items[0];
                            let file = attachment;
                            file.Load().then(() => {
                                //console.log(file.Base64Content);
                                if ((c.AttachmentFilter == null || c.AttachmentFilter == "") || (c.AttachmentFilter != null && file.Name.indexOf(c.AttachmentFilter) !== -1)) {
                                    Alteryx.Engine.SendMessage.Info(file.Name);
                                    Alteryx.Engine.SendMessage.Info(Buffer(file.Base64Content,"base64").toString());
                                    // Write out to Alteryx output at this point either as file stream or file save.
                                }
                            }, (error) => {
                                if (error) {
                                    Alteryx.Engine.SendMessage.Error(error)
                                }
                            });
                        }
                    }
                }
                Alteryx.Engine.SendMessage.Info(record);
            }, (error) => {
                if (error) {
                    Alteryx.Engine.SendMessage.Error(error);
                    Alteryx.Engine.SendMessage.Complete();
                }
            });
        }

        //console.log('Results: ' + Array.from(results.Items).length);
        //console.log('Page Offset: ' + itemView.Offset);
        //console.log('More Records Available? ' + results.MoreAvailable);
        if (results.MoreAvailable === true) {
            itemView.Offset += Array.from(results.Items).length;

            GetItemsFromFolder(service, c.SkipRootFolder ? folder.Id : folder, itemView);
        }
        
        Alteryx.Engine.SendMessage.Complete();
    }, function (errors) {
        Alteryx.Engine.SendMessage.Error(errors);
        Alteryx.Engine.SendMessage.Complete();
    });*/
    
    if (folder == 0 && c.IncludeRecurringEvents) {
        //console.log('Recurring Events Section Reached.');
        service.FindAppointments(folder, calendarView)
        .then(function (results) {
            for (var item of results.Items) {
                EwsJS.Item.Bind(service, item.Id, new EwsJS.PropertySet(EwsJS.BasePropertySet.IdOnly, c.Fields))
                .then(function (boundItem) {
                    var record = [];
                    for (var field of c.Fields) {
                        // Replace with Alteryx Engine message record output.
                        if (field.Name === 'Id') {
                            record.push(boundItem.propertyBag.properties.objects[field.Name].UniqueId);
                        } else {
                            if (field.Name != 'Attachments') {
                                record.push(boundItem.propertyBag.properties.objects[field.Name].toString());
                            }
                        }
    
                        // Replace with Alteryx Engine attachment record output.
                        if (field.Name === 'Attachments' && boundItem.HasAttachments) {
                            for (var attachment of boundItem.Attachments.items) {
                                let file = attachment;
                                file.Load().then(() => {
                                    //console.log(file.Base64Content);
                                    Alteryx.Engine.SendMessage.Info(file.Name);
                                    Alteryx.Engine.SendMessage.Info(Buffer(file.Base64Content,"base64").toString());
                                    // Write out to Alteryx output at this point either as file stream or file save.
                                }, (error) => {
                                    if (error) {
                                        Alteryx.Engine.SendMessage.Error(error);
                                    }
                                });
                            }
                        }
                    }
                    Alteryx.Engine.SendMessage.Info(record);
                }, (error) => {
                    if (error) {
                        Alteryx.Engine.SendMessage.Error(error);
                        Alteryx.Engine.SendMessage.Complete();
                    }
                });
            }
    
            //console.log('Results: ' + Array.from(results.Items).length);
            //console.log('Page Offset: ' + itemView.Offset);
            //console.log('More Records Available? ' + results.MoreAvailable);
            if (results.MoreAvailable === true) {
                itemView.Offset += Array.from(results.Items).length;
    
                GetItemsFromFolder(service, folder, itemView);
            }
        }, function (errors) {
            Alteryx.Engine.SendMessage.Error(errors);
            Alteryx.Engine.SendMessage.Complete();
        });

        // This completion kills the bound item iteration output. 
        //Alteryx.Engine.SendMessage.Complete();
    }
}

// Get folders for UI drop-downs.
function getWellKnownFolderName(index) {
    var folders = Object.keys(EwsJS.WellKnownFolderName).map(function(folder) {
        return EwsJS.WellKnownFolderName[folder];
    });
    // Inbox.
    return folders[index];
}

// Get server versions for UI drop-downs.
function getExchangeVersion(index) {
    var exchangeVersions = Object.keys(EwsJS.ExchangeVersion).map(function(version) {
        return EwsJS.ExchangeVersion[version];
    });
    // 7 = Exchange 2016.
    return exchangeVersions[index];
}

// Get items from Inbox.
//GetItemsFromFolder(service, 4, c.QueryString, true);

// Get items from Calendar.
//GetItemsFromFolder(service, 0, c.QueryString, c.isRoot);