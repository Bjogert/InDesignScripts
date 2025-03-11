#target "InDesign"

(function(){
    app.doScript(main, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Flexible Replace Labels and Contents");

    function main(){
        if(!app.documents.length){
            alert("No document open."); return;
        }

        var doc = app.activeDocument;
        var priceGroups=["SEK_SE","SEK_NO","DKK_DK","EUR_EU","GBP_UK","EUR_BEIR","EUR_EURAEA","USD_US","EUR_DE"];

        var initialLabel="";

        if(app.selection.length===1 && app.selection[0] instanceof TextFrame){
            initialLabel=app.selection[0].label;
            for(var p=0;p<priceGroups.length;p++){
                var suffix="_"+priceGroups[p];
                if(initialLabel.substr(-suffix.length)===suffix){
                    initialLabel=initialLabel.substr(0,initialLabel.length-suffix.length);
                    break;
                }
            }
        }else{
            initialLabel="";
        }

        var dlg=app.dialogs.add({name:"Replace Script Labels",canCancel:true});
        var col1=dlg.dialogColumns.add();

        var oldLabelField=col1.textEditboxes.add({staticLabel:"Old base label:",editContents:initialLabel,minWidth:180});
        var newLabelField=col1.textEditboxes.add({staticLabel:"New base label:",editContents:"",minWidth:180});

        var replaceContentCheckbox=col1.checkboxControls.add({
            staticLabel:"Replace frame contents",
            checkedState:true
        });

        var col2=dlg.dialogColumns.add();
        col2.staticTexts.add({staticLabel:"Select Price Groups:"});
        var gpCol=col2.borderPanels.add().dialogColumns.add();

        var cbxGroups=[];
        for(var i=0;i<priceGroups.length;i++){
            cbxGroups.push(gpCol.checkboxControls.add({staticLabel:priceGroups[i],checkedState:false}));
        }
        var cbAll=gpCol.checkboxControls.add({staticLabel:"ALL price groups",checkedState:false});

        if(!dlg.show()){dlg.destroy();return;}

        var oldBase=oldLabelField.editContents;
        var newBase=newLabelField.editContents;
        var replaceContent=replaceContentCheckbox.checkedState;

        var selectedGroups=cbAll.checkedState?priceGroups.slice():[];
        if(!cbAll.checkedState){
            for(var i=0;i<cbxGroups.length;i++){
                if(cbxGroups[i].checkedState)selectedGroups.push(priceGroups[i]);
            }
        }

        dlg.destroy();

        if(!oldBase||!newBase||selectedGroups.length===0){
            alert("Please provide necessary details.");return;
        }

        var layers=doc.layers,states=[];
        for(var i=0;i<layers.length;i++){
            states[i]={locked:layers[i].locked,visible:layers[i].visible};
            layers[i].locked=false;layers[i].visible=true;
        }

        var frameMap={};
        var frames=doc.textFrames.everyItem().getElements();
        for(var f=0;f<frames.length;f++){
            frameMap[frames[f].label]=frames[f].contents;
        }

        var replacedCount=0,pagesAffected={},changedFrames=[];

        for(var f=0;f<frames.length;f++){
            var frame=frames[f];
            for(var g=0;g<selectedGroups.length;g++){
                var oldFullLabel=oldBase+"_"+selectedGroups[g];
                var newFullLabel=newBase+"_"+selectedGroups[g];

                if(frame.label===oldFullLabel){
                    frame.label=newFullLabel;

                    if(replaceContent){
                        var replacementContent=frameMap[newFullLabel];
                        if(replacementContent==null){
                            replacementContent=prompt("Enter content for "+newFullLabel+":","New content here");
                            if(replacementContent==null)continue;
                            frameMap[newFullLabel]=replacementContent;
                        }
                        frame.contents=replacementContent;
                    }

                    replacedCount++;
                    pagesAffected[frame.parentPage.name]=true;
                    changedFrames.push(frame);
                }
            }
        }

        for(var i=0;i<layers.length;i++){
            layers[i].locked=states[i].locked;
            layers[i].visible=states[i].visible;
        }

        var affectedPagesList=[];
        for(var page in pagesAffected){
            if(pagesAffected.hasOwnProperty(page))affectedPagesList.push(page);
        }
        affectedPagesList.sort(function(a,b){return Number(a)-Number(b);});

        if(changedFrames.length>0){
            app.select(changedFrames);
            alert("Replaced "+replacedCount+" labels.\nPages: "+affectedPagesList.join(", ")+"\n\nBoxes selected.");
        }else{
            alert("No matching labels found.");
        }
    }
})();
