#!/usr/bin/env node 
var g_winax = require('./activex');
var path = require('path');

if(process.argv.length<3)
{
    console.error("Missing argument. Usage:\nnodewscript SomeJS.js <other arguments>")
}

// For the case we want to use 'require' or 'module' - we need to make it available through global.
global.require = require;
global.exports = exports;
global.module = module;

global.SeSIncludeFile = function (includePath, isLocal, ctx) {

    var vm = require("vm");

    var fs = require("fs");
    var path = require("path");

    var absolutePath = path.resolve(includePath);
    var data = fs.readFileSync(absolutePath);
    var script = new vm.Script(data, { filename: absolutePath, displayErrors: true });
    if (isLocal && ctx) {
        if (!ctx.__runctx__) ctx.__runctx__ = vm.createContext(ctx);
        script.runInContext(ctx.__runctx__);
    } else {
        script.runInThisContext();
    }

    return '';
};

global.GetObject = function GetObject(strMoniker) {
    return new g_winax.Object(strMoniker, {getobject:true});
}

var WinAXActiveXObject = ActiveXObject;

function ActiveXObjectCreate(progId, pfx)
{
    var res = new WinAXActiveXObject(progId);
    if(progId=="SeSHelper")
    {
        // SeSHelper implements Include that we want to override
        res = new Proxy({helper:res}, {
            get(target, propKey, receiver) {
                return function (...args) {
                    var helper = target.helper;
                    var ovr =
                    {
                        IncludeLocal(/**Object*/fName, ctx) /**Object*/
                        {
                            return this.Include(fName, true, ctx);
                        },
                
                        Include(/**Object*/fName, local, ctx) /**Object*/
                        {
                            var includePath = helper.ResolvePath(fName);
                            if(!includePath)
                            {
                                var err = "Error trying to include file. File not found:"+fName;
                                return "console.error("+JSON.stringify(err)+")";
                            }
                        
                            if(!local) 
                            {
                                helper.SetAlreadyIncluded(includePath);
                            }
                    
                            return global.SeSIncludeFile(includePath, local, ctx);
                    
                        },
                
                        IncludeOnce(/**Object*/fName) /**Object*/
                        {
                            if(!helper.IsAlreadyIncluded(fName))
                            {
                                return this.Include(fName);
                            }
                            return "";
                        }
                    }

                    if(propKey in ovr)
                    {
                        return ovr[propKey](...args);
                    } else {
                        let result = target.helper[propKey](...args);
                        return result;    
                    }

                };
            }
          });
    }

    if(res&&pfx)
    {
        WScript.ConnectObject(res, pfx);
    }
    
    return res;
}

global.ActiveXObject=new Proxy(ActiveXObject, {
    construct(target,args)
    {
        //console.log("Creating ActiveXObject: "+args);
        return ActiveXObjectCreate(...args);
    }
});

global.Enumerator = function (arr) {
    this.arr = arr;
    this.enum = arr._NewEnum;
    this.nextItem = null;
    this.atEnd = function () { return this.nextItem===undefined; }
    this.moveNext = function () {  this.nextItem = this.enum.Next(1); return this.atEnd(); }
    this.moveFirst = function () { this.enum.Reset(); this.nextItem=null; return this.moveNext(); }
    this.item = function () { return this.nextItem; }
    
    this.moveNext();
}

// See interface here
// https://www.devguru.com/content/technologies/wsh/objects-wscript.html

var WScript = {
    // Properties
    Application : null,
    BuildVersion : "6.0",
    FullName : null,
    Interactive : true,
    Name : "Windows / Node Script Host",
    Path : __filename,
    ScriptFullName : null,
    ScriptName : null,
    StdErr: {
        Write(txt) { console.log(txt); }
    },
    StdIn: {
        ReadLine() { 
            var readline = require('readline');
            var gotLine = false;
            var line = undefined;

            var rl = readline.createInterface({
              input: process.stdin,
              output: process.stdout,
              terminal: false
            });
            
            rl.on('line', function (cmd) {
                line = cmd;
                rl.close();
            });
            while(!gotLine)
            {
                WScript.Sleep(10);
            }
            return line
        }
    },
    StdOut: {
        Write(txt) { console.err(txt); }
    },
    Timeout : -1,
    Version : 'NODE.WIN32',

    // Methods
    
    Echo : console.log,

    __Connected : [],

    ConnectObject(obj, pfx)
    {
        var connectionPoints = g_winax.getConnectionPoints(obj);
        var connectionPoint = connectionPoints[0];
        var allMethods = connectionPoint.getMethods();

        var advobj={};
        var found = false;
        for(var i=0;i<allMethods.length;i++)
        {
            var a = allMethods[i];
            
            var cpFunc = null;
            var fName = pfx + a;
            
            eval('if(typeof '+fName+'!="undefined") cpFunc = '+fName+';')

            if(cpFunc)
            {
                found = true;
                advobj[a]=cpFunc;
            }
        }

        if(found)
        {
            var dwCookie = connectionPoint.advise(advobj);
            this.__Connected.push({obj,pfx,connectionPoint,dwCookie});
        }

    },

    DisconnectObject(obj)
    {
        for(var i=0;i<this.__Connected.length;i++)
        {
            var o = this.__Connected[i];
            if(o.obj==obj)
            {
                o.connectionPoint.unadvise(o.dwCookie);
                // Remove connection from list
                this.__Connected.splice(i,1);
                break;
            }
        }
    },
    
    CreateObject : ActiveXObjectCreate,

    GetObject : ActiveXObjectCreate,

    Quit (exitCode) { 
        // Gracefully disconnect and release all ActiveX objects, if needed
        while(this.__Connected.length>0)
        {
            WScript.DisconnectObject(this.__Connected[0].obj);
        }
        WScript.Sleep(60);
        process.exit(exitCode); 
    },

    Sleep (ms) { 
        var begin = (new Date()).valueOf();
        MessageLoop();
        var dt = (new Date()).valueOf()-begin;

        while(dt<ms)
        {
            MessageLoop();
            g_winax.winaxsleep(1);
            dt = (new Date()).valueOf()-begin;
        }
    },

    __Args : process.argv,
    __ShowLogo : true,
    __Debug : false,

    // Collections
    Arguments : null,

    __InitArgs(){
        this.Application = this;
        this.__Args=[];
        for(let i=0;i<process.argv.length;i++)
        {
            if(i==0)
            {
                this.FullName = path.resolve(process.argv[0])
            } else if(i==1) {
                this.Name = process.argv[1];
            } else if(i==2) {
                this.ScriptFullName = path.resolve(process.argv[2]);
                this.ScriptName = path.basename(this.ScriptFullName);
            } else {
                // All other args
                var arg = process.argv[i];
                if(arg.startsWith("//"))
                {
/* 
Parse all arguments and process those related to core WScript:
 //B         Batch mode: Suppresses script errors and prompts from displaying
 //D         Enable Active Debugging
 //E:engine  Use engine for executing script
 //H:CScript Changes the default script host to CScript.exe
 //H:WScript Changes the default script host to WScript.exe (default)
 //I         Interactive mode (default, opposite of //B)
 //Job:xxxx  Execute a WSF job
 //Logo      Display logo (default)
 //Nologo    Prevent logo display: No banner will be shown at execution time
 //S         Save current command line options for this user
 //T:nn      Time out in seconds:  Maximum time a script is permitted to run
 //X         Execute script in debugger
 //U         Use Unicode for redirected I/O from the console
*/
                    if(arg.toLowerCase()=="//nologo")
                    {
                        this.__ShowLogo = false;
                    } else if(arg.toLowerCase().indexOf("//t:")==0) {
                        var timeOut = parseInt(arg.substr("//t:".length));
                        this.Timeout = timeOut;
                        setTimeout(()=>{WScript.Quit(0);},timeOut*1000)
                    } else if(arg.toLowerCase()=="//x"||arg.toLowerCase()=="//d") {
                        this.__Debug = true;
                    }
                } else {
                    this.__Args.push(arg);
                }
            }
        }

        // The trick starts here
        // We want WScript.Arguments to be accessible in all of the following ways:
        // WScript.Arguments(i)
        // WScript.Arguments.item(i)
        // WScript.Arguments.length
        // Since it is not directly possible, we use a number of tricks.
        // Function.length is a number of expected arguments
        // And we make this 'fake' number
        var argArr=[];
        for(var ai=0;ai<WScript.__Args.length;ai++)argArr.push("arg"+ai);
        var ArgF=null;
        eval("ArgF = function ("+argArr.join(",")+") { return WScript.__Args[arguments[0]];  }");
        ArgF.item = function(i) { return WScript.__Args[i] };
        ArgF.toString = function() { return WScript.__Args.join(' '); }
        ArgF.Count = function() { return WScript.__Args.length; };

        this.Arguments = ArgF;
    }
}

WScript.__InitArgs();

global.WScript = WScript;

if(WScript.__ShowLogo)
{
    WScript.Echo(WScript.Name+" "+WScript.BuildVersion);
}

function MessageLoop()
{
    g_winax.peekAndDispatchMessages(); // allows ActiveX event to be dispatched
}

const __rapise_global_context__ = require("vm").createContext(this);
global.__rapise_global_context__ = __rapise_global_context__;

if(WScript.ScriptFullName)
{
    SeSIncludeFile(WScript.ScriptFullName);
} else {
    WScript.StdErr.Write("File.js missing.\nUsage wscriptnode <File.js>")
}
