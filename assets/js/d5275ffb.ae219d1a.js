"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[5580],{3905:function(e,t,o){o.d(t,{Zo:function(){return h},kt:function(){return p}});var a=o(7294);function n(e,t,o){return t in e?Object.defineProperty(e,t,{value:o,enumerable:!0,configurable:!0,writable:!0}):e[t]=o,e}function r(e,t){var o=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),o.push.apply(o,a)}return o}function i(e){for(var t=1;t<arguments.length;t++){var o=null!=arguments[t]?arguments[t]:{};t%2?r(Object(o),!0).forEach((function(t){n(e,t,o[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(o)):r(Object(o)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(o,t))}))}return e}function s(e,t){if(null==e)return{};var o,a,n=function(e,t){if(null==e)return{};var o,a,n={},r=Object.keys(e);for(a=0;a<r.length;a++)o=r[a],t.indexOf(o)>=0||(n[o]=e[o]);return n}(e,t);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);for(a=0;a<r.length;a++)o=r[a],t.indexOf(o)>=0||Object.prototype.propertyIsEnumerable.call(e,o)&&(n[o]=e[o])}return n}var l=a.createContext({}),u=function(e){var t=a.useContext(l),o=t;return e&&(o="function"==typeof e?e(t):i(i({},t),e)),o},h=function(e){var t=u(e.components);return a.createElement(l.Provider,{value:t},e.children)},c={inlineCode:"code",wrapper:function(e){var t=e.children;return a.createElement(a.Fragment,{},t)}},d=a.forwardRef((function(e,t){var o=e.components,n=e.mdxType,r=e.originalType,l=e.parentName,h=s(e,["components","mdxType","originalType","parentName"]),d=u(o),p=n,m=d["".concat(l,".").concat(p)]||d[p]||c[p]||r;return o?a.createElement(m,i(i({ref:t},h),{},{components:o})):a.createElement(m,i({ref:t},h))}));function p(e,t){var o=arguments,n=t&&t.mdxType;if("string"==typeof e||n){var r=o.length,i=new Array(r);i[0]=d;var s={};for(var l in t)hasOwnProperty.call(t,l)&&(s[l]=t[l]);s.originalType=e,s.mdxType="string"==typeof e?e:n,i[1]=s;for(var u=2;u<r;u++)i[u]=o[u];return a.createElement.apply(null,i)}return a.createElement.apply(null,o)}d.displayName="MDXCreateElement"},9254:function(e,t,o){o.r(t),o.d(t,{frontMatter:function(){return s},contentTitle:function(){return l},metadata:function(){return u},toc:function(){return h},default:function(){return d}});var a=o(7462),n=o(3366),r=(o(7294),o(3905)),i=["components"],s={},l=void 0,u={unversionedId:"overview-of-adv-processing/10-01 some-thoughts-on-creating-a-workflow",id:"overview-of-adv-processing/10-01 some-thoughts-on-creating-a-workflow",title:"10-01 some-thoughts-on-creating-a-workflow",description:"some thoughts on creating a workflow",source:"@site/docs/10-overview-of-adv-processing/10-01 some-thoughts-on-creating-a-workflow.md",sourceDirName:"10-overview-of-adv-processing",slug:"/overview-of-adv-processing/10-01 some-thoughts-on-creating-a-workflow",permalink:"/excel-vba-book/docs/overview-of-adv-processing/10-01 some-thoughts-on-creating-a-workflow",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/10-overview-of-adv-processing/10-01 some-thoughts-on-creating-a-workflow.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"overview of adv processing",permalink:"/excel-vba-book/docs/overview-of-adv-processing/10 overview-of-adv-processing"},next:{title:"overview of events",permalink:"/excel-vba-book/docs/overview-of-events/11 overview-of-events"}},h=[{value:"some thoughts on creating a workflow",id:"some-thoughts-on-creating-a-workflow",children:[{value:"inputs",id:"inputs",children:[],level:3},{value:"outputs",id:"outputs",children:[],level:3},{value:"intermediate results",id:"intermediate-results",children:[],level:3},{value:"putting it all together",id:"putting-it-all-together",children:[],level:3}],level:2}],c={toc:h};function d(e){var t=e.components,o=(0,n.Z)(e,i);return(0,r.kt)("wrapper",(0,a.Z)({},c,o,{components:t,mdxType:"MDXLayout"}),(0,r.kt)("h2",{id:"some-thoughts-on-creating-a-workflow"},"some thoughts on creating a workflow"),(0,r.kt)("p",null,"If you are sitting down to create an advanced workflow, there are a handful of things to consider. The list that follows is not complete nor is it meant to include items that are always relevant. The problem with these lists is that with a general programming environment like Excel, it's impossible to describe everything to consider. Having said that, I have built tons of these workflows and can comment on a handful of things that nearly always come up. The first item to touch on is the general structure/outline of a VBA workflow. This breakdown seems to always hold true."),(0,r.kt)("p",null,"Your VBA workflow will contain steps or sub steps that roughly be described as:"),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"Inputs"),(0,r.kt)("li",{parentName:"ul"},"Intermediate results"),(0,r.kt)("li",{parentName:"ul"},"Outputs")),(0,r.kt)("p",null,"If your workflow is advanced enough to include a number of sub steps built from other steps, then you are likely to find that this breakdown applies within and across levels of your workflow. That is, the outputs of one step may very well be the inputs to another step. The intermediate result from one action will be the input for another."),(0,r.kt)("p",null,"When thinking in terms of these categories, there is a useful distinction to make that is somewhat unique to Excel programming: do your inputs and outputs exist in the Excel spreadsheet or only in the VBA code? This distinction is meaningful because it helps you think about how much of your workflow is the automation of otherwise human tasks (which could still be done by a human) vs. steps that are purely programmatic and could not be replicated by a human. Where this distinction is most likely to show up is when you are deciding where and how to perform a calculations. In theory, all of the Excel spreadsheet could be done in VBA via the ",(0,r.kt)("inlineCode",{parentName:"p"},"WorksheetFunction")," object. Doing everything in VBA defeats a large part of the benefit that comes from programming with VBA. It's easy to lose sight of this when you see a clean code-only solution to a problem, but realize that the greatest benefit to programming alongside Excel is that you have a powerful, human readable scratch pad that lives alongside your VBA."),(0,r.kt)("p",null,"As a comment, I have seen incredibly complicated workflows that involved detailed calculations of arrays that were donely exclusively in VBA. The math was fine and the results were generally useful. The problem was that there was no way to spot check a given result without debugging code. This makes it nearly impossibly for someone without VBA experience to validate your work. It also provides you job security, but ideally you'd gain security by other means."),(0,r.kt)("p",null,"A better marriage of VBA and Excel is to utilize Excel for all of the tasks it's great at: calculation, visual outputs, charting, page layouts and printing, and also the deep data oriented features (sorting, filtering, etc). Where VBA comes in handy, is wiring together all of these items into a coherent package that runs more efficiently than anything that a human alone could do. The best workflows typically take a very simply underlying spreadsheet and apply to a large number of items. In this way, you are able to spot check a single result, verify the formulas, and investigate an interesting result. You are also free to just hit go and have 10,000+ realists streamed into a table for consumption. IF you find yourself looking for all sorts of tricks to avoid using the underlying Excel model for your programming, I'd strongly encourage to just switch to a fully programmatic language that does not have the Excel UI. You will save yourself a ton of headache. If you are only aware of VBA and looking to push the envelope in terms of performance, then that's an OK place to be. Just realize that there are better alternatives to Excel for high performance computing."),(0,r.kt)("h3",{id:"inputs"},"inputs"),(0,r.kt)("p",null,'Back to the overall structure, there are inputs, outputs, and intermediate results. Depending on what you are doing, some of these aspects may just exist on/within the spreadsheet and be easy to overlook as an input or output. It\'s not until you wire up a more complicated workflow that you are forced to recognize the different pieces in a spreadsheet for what they are. On the input front, there are a handful of items that should trigger your thought of "input":'),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"A file that contains some data to be processed, filtered, etc."),(0,r.kt)("li",{parentName:"ul"},"A couple of columns in a spreadsheet that need to be processed and then charted."),(0,r.kt)("li",{parentName:"ul"},"15 scattered cells that meet some criteria within a block of data"),(0,r.kt)("li",{parentName:"ul"},"THe contents of the clipboard from another program"),(0,r.kt)("li",{parentName:"ul"},"The formatting of a couple of cells")),(0,r.kt)("p",null,"All of those items could be used as the input to a VBA workflow. Some of these items are odd to think about if you are coming from a noter programming environment. What does it mean for the formatting of a cell to be an input? Well Excel provides you with a rich Object Model full of metadata about all of the various cells of data. That metadata can be as useful as actually structured data if there is a structure to it. I've seen it countless times where someone has methodically bolded all of the cells of intereste in a block of data. That bold format is as good as some field called ",(0,r.kt)("inlineCode",{parentName:"p"},"Important = True")," which could then be processed in another language. Instead of that flag, you just check ",(0,r.kt)("inlineCode",{parentName:"p"},"Range.Format.Bold = True"),". This of course relies on an implicit assumption about how the data is structured, but this is common in the Excel/VBA world."),(0,r.kt)("p",null,"Excel also has a very strong UI which makes it possible to immediately solicit user input in a way that is not easily replicated coming from other languages. Where this shows up most frequently is when you start using the ",(0,r.kt)("inlineCode",{parentName:"p"},"ActiveCell"),", ",(0,r.kt)("inlineCode",{parentName:"p"},"ActiveWorkbook"),", ",(0,r.kt)("inlineCode",{parentName:"p"},"Selection"),' and other objects which are dependent on user input. In a lot of other languages you have to spend a ton of time pointing the program to the correct file, or rows, or columns, or other items to process. In Excel, you leverage the fact that most people know how to select or activate items they want, and you can use that user input as an actual input to your VBA. This becomes quite powerful when you are building utility code that may be used across multiple workbooks. This becomes much harder in other languages where the idea of a "open file" is far less well defined. You certainly cannot query the selected cells in an R data table.'),(0,r.kt)("h3",{id:"outputs"},"outputs"),(0,r.kt)("p",null,"THe next item to hit are the outpost of a workflow. Very often, the outputs are obvious because you had some task to complete with VBA, and the outputs are simply the results of that task. Where things become more complicated is when you string together steps and the output of one becomes the input for the next. When that happens, you often have to decided what intermediate format is best for the transfer. You may or may not settle on a format that is easily human consumable. There are tradeoffs here that will be discussed later. The output of a workflow can be a number of things:"),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"A string, number, cell, row, column, or table of data that was processed by the VBA"),(0,r.kt)("li",{parentName:"ul"},"A chart"),(0,r.kt)("li",{parentName:"ul"},"A collection of shapes"),(0,r.kt)("li",{parentName:"ul"},"A worksheet that includes any of the items above"),(0,r.kt)("li",{parentName:"ul"},"A workbook that includes a number of constructed worksheets"),(0,r.kt)("li",{parentName:"ul"},"A change to the formatting of a number of cells"),(0,r.kt)("li",{parentName:"ul"},"A change to the properties of a Range, Worksheet or Workbook"),(0,r.kt)("li",{parentName:"ul"},"A new text file written to disk"),(0,r.kt)("li",{parentName:"ul"},"Some result output to the Clipboard"),(0,r.kt)("li",{parentName:"ul"},"Pages of physical paper if your VBA prints"),(0,r.kt)("li",{parentName:"ul"},"Some change to the filesystem or disk"),(0,r.kt)("li",{parentName:"ul"},"Some other program opened or run with specific parameters")),(0,r.kt)("p",null,"This is a shortened list since the possibilities here are closer to endless. The idea however is that you can effect a large amount of change from VBA and so your possible outputs can be quite numerous. A typically workflow will accumulate a large number of these outputs individually and will then produce some final product which highlights some of those outputs."),(0,r.kt)("h3",{id:"intermediate-results"},"intermediate results"),(0,r.kt)("p",null,"When discussing intermediate results, it is generally best to limit your thoughts to whatever will live only in VBA. In that sense, the question of intermediate results is: what programming constructs can exist without the user ever seeing them? Sometimes you need to determine the unique items in a list to do some processing. Do you generate that list of unique items in Excel somewhere? Or, do you determine the unique items using VBA and then output some result which may or may not include the full list of unique items. If you are doing the former, Excel provides a nice ",(0,r.kt)("inlineCode",{parentName:"p"},"RemoveDuplicates")," function which will replicate the ",(0,r.kt)("inlineCode",{parentName:"p"},"Data->Remove Duplicates")," functionality. This works great if you want the user to see the final list of values. You can also use a ",(0,r.kt)("inlineCode",{parentName:"p"},"Dictionary")," in VBA to only store the unique vaults from a list. In this sense, the ",(0,r.kt)("inlineCode",{parentName:"p"},"Dictionary")," represents an intermediate value that may not be shown to the user. You will make this decision several times before you realize that you are deciding whether or not something should exist in VBA only. Often times, the decision does not matter, but for certain workflows it can make a huge difference."),(0,r.kt)("p",null,"An example is a multi step process where you might want the user to verify the calculations so far and correct any errors. This can technically be done with VBA or Excel, but it is much easier to ask a user to verify an Excel spreadsheet than to debug the code and check ",(0,r.kt)("inlineCode",{parentName:"p"},"Locals"),". If you need to do this verification step, then it makes a lot of sense to use an intermediate result that dumps back into Excel. In this sense, you've taken an intermediate result and converted it to an output. That output may or may not be modified by the user and it then becomes the input for the next step."),(0,r.kt)("h3",{id:"putting-it-all-together"},"putting it all together"),(0,r.kt)("p",null,"Having given a snapshot of the options for inputs and outputs, it's worth commenting generally on how they all fit together. Your goal should be to build a workflow that consists of steps that can all be described individually and possibly run on their own. Your task is then generating these individual steps and determining how to wire them together. The most common approach to building these workflows is that you start with some single task and then the scope expands as the analysis expands. You can build the ultimate workhorse of a workflow initially, or you can adapt your code to the task as the task comes into view. Depending on where you're starting and the definition at the start, you will determine how complicated to make things at the start."),(0,r.kt)("p",null,"It is very common to start with a single, straight-through workflow and then build it out into Modules as the work expands. IN this way, you are constantly reevaluating the inputs nad output sof your program to build the smaller blocks which need these definitions. In my experience, nearly all VBA workflows will take shape in this process eventually. It's quite rate to build a complicate workflow once and for all. Generally you start simple and end up with a full featured application at the end."))}d.isMDXComponent=!0}}]);