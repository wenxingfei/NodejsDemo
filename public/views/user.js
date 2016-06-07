/**
 * Created by wWX247609 on 2016/4/12.
 */
define([
],function(){

    var buttons = {
        cols:[
            {},
            {
                view:"button",label:"Add",click:function(){
                    $$("user_grid").add({},0);
                }
            },
            {
                view:"button",label:"Export",click:function(){
                    //webix.ajax().get("/api/users/export5", { }, function(path, xml, xhr){
                    //    //response
                    //    debugger;
                    //    path = "http://localhost.huawei.com:5000/temp/" + path;
                    //    window.open(path);
                    //});
                    window.location = '/api/users/export5';
                    //webix.toExcel($$("user_grid"), {
                    //    columns:{
                    //        "username":{header: "User Name", width: 200, template: webix.template("#id#.#username#")},
                    //        "email":{header: "Email", width: 100},
                    //        "type":{header: "Type", width: 100},
                    //    }
                    //});
                }
            }
        ]
    };

    var grid = {
        view:"datatable",
        editable:true,
        id:"user_grid",
        editaction:"dblclick",
        columns:[
            { id:"username",    header:"username",fillspace:true, editor:"text"},
            { id:"email",   header:"email", width:300, editor:"text"},
            { id:"type",    header:"type", width:100, editor:"text"}
        ],
        url:"rest->/api/users",
        save:{
            autoupdate :true,
            url:"rest->/api/users"
        }
    };

    var ui = {
        rows:[
            buttons,
            grid
        ]
    }

    return {
        $ui: ui,
        //$oninit:function(view){
        //    view.parse(records.data);
        //}
    };

});