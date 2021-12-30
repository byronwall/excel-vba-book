"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[9741],{3905:function(e,t,n){n.d(t,{Zo:function(){return c},kt:function(){return d}});var a=n(7294);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function i(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,a)}return n}function r(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?i(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):i(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function l(e,t){if(null==e)return{};var n,a,o=function(e,t){if(null==e)return{};var n,a,o={},i=Object.keys(e);for(a=0;a<i.length;a++)n=i[a],t.indexOf(n)>=0||(o[n]=e[n]);return o}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(a=0;a<i.length;a++)n=i[a],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(o[n]=e[n])}return o}var h=a.createContext({}),s=function(e){var t=a.useContext(h),n=t;return e&&(n="function"==typeof e?e(t):r(r({},t),e)),n},c=function(e){var t=s(e.components);return a.createElement(h.Provider,{value:t},e.children)},p={inlineCode:"code",wrapper:function(e){var t=e.children;return a.createElement(a.Fragment,{},t)}},u=a.forwardRef((function(e,t){var n=e.components,o=e.mdxType,i=e.originalType,h=e.parentName,c=l(e,["components","mdxType","originalType","parentName"]),u=s(n),d=o,m=u["".concat(h,".").concat(d)]||u[d]||p[d]||i;return n?a.createElement(m,r(r({ref:t},c),{},{components:n})):a.createElement(m,r({ref:t},c))}));function d(e,t){var n=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var i=n.length,r=new Array(i);r[0]=u;var l={};for(var h in t)hasOwnProperty.call(t,h)&&(l[h]=t[h]);l.originalType=e,l.mdxType="string"==typeof e?e:o,r[1]=l;for(var s=2;s<i;s++)r[s]=n[s];return a.createElement.apply(null,r)}return a.createElement.apply(null,n)}u.displayName="MDXCreateElement"},1117:function(e,t,n){n.r(t),n.d(t,{frontMatter:function(){return l},contentTitle:function(){return h},metadata:function(){return s},toc:function(){return c},default:function(){return u}});var a=n(7462),o=n(3366),i=(n(7294),n(3905)),r=["components"],l={},h=void 0,s={unversionedId:"overview-of-charting/06-01 introduction-to-charting",id:"overview-of-charting/06-01 introduction-to-charting",title:"06-01 introduction-to-charting",description:"introduction to charting",source:"@site/docs/06-overview-of-charting/06-01 introduction-to-charting.md",sourceDirName:"06-overview-of-charting",slug:"/overview-of-charting/06-01 introduction-to-charting",permalink:"/docs/overview-of-charting/06-01 introduction-to-charting",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/06-overview-of-charting/06-01 introduction-to-charting.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"overview of charting",permalink:"/docs/overview-of-charting/06 overview-of-charting"},next:{title:"06-02 common-objectsproperties-for-a-Chart",permalink:"/docs/overview-of-charting/06-02 common-objectsproperties-for-a-Chart"}},c=[{value:"introduction to charting",id:"introduction-to-charting",children:[{value:"a quick overview of the object model",id:"a-quick-overview-of-the-object-model",children:[],level:3},{value:"obtaining a reference to a Chart",id:"obtaining-a-reference-to-a-chart",children:[{value:"<code>ActiveChart</code>",id:"activechart",children:[],level:4},{value:"<code>Selection</code>",id:"selection",children:[],level:4},{value:"ChartObjects",id:"chartobjects",children:[],level:4},{value:"Workbook.Sheets to get Chart references",id:"workbooksheets-to-get-chart-references",children:[],level:4}],level:3}],level:2}],p={toc:c};function u(e){var t=e.components,n=(0,o.Z)(e,r);return(0,i.kt)("wrapper",(0,a.Z)({},p,n,{components:t,mdxType:"MDXLayout"}),(0,i.kt)("h2",{id:"introduction-to-charting"},"introduction to charting"),(0,i.kt)("p",null,"Charting is the second most important aspect of automatic Excel behind manipulating ",(0,i.kt)("inlineCode",{parentName:"p"},"Ranges"),'. There is a bias when saying that because a lot of what I do after engineering calculations is chart the results. In particular, Excel can be used to great effect to chart time series of data. The other reason charts are so amenable to VBA is that very often you are applying the same actions to the charts. In that sense, the VBA related to charts is doing a lot of changing settings and formats so that the charts look the way you want. This ahs the immediate effect of making your charts look less like "they came from Excel" which is a common knock in some circles.'),(0,i.kt)("p",null,"When working with ",(0,i.kt)("inlineCode",{parentName:"p"},"Charts"),", there is a ",(0,i.kt)("inlineCode",{parentName:"p"},"Range")," of difficulties depending on what you are trying to do. In some cases, working with an existing ",(0,i.kt)("inlineCode",{parentName:"p"},"chart")," is much easier than creating a new one. In other instances, it can be much simpler to create a new chart, starting from a default, rather than change all the settings back. One other major difference between ",(0,i.kt)("inlineCode",{parentName:"p"},"Charts")," and ",(0,i.kt)("inlineCode",{parentName:"p"},"Ranges")," is that working with charts is much more about knowing the object model than knowing how to program. The vast majority of your code related to charts is simply iterating through objects to find the one property that you want to change. This makes it easier to write chart VBA once you have the basics of ",(0,i.kt)("inlineCode",{parentName:"p"},"For Each")," loops down. It also means that you need to spend some time getting comfortable with the object model."),(0,i.kt)("p",null,"There is one oddity related to Charts that is worth mentioning now. Charts can either be embedded as an object on a ",(0,i.kt)("inlineCode",{parentName:"p"},"Worksheet"),", or they can be their own ",(0,i.kt)("inlineCode",{parentName:"p"},"Sheets"),". I personally never use the latter case, but it is common enough that it needs to be on your mind when working with Charting code."),(0,i.kt)("p",null,"(I don't use the Chart as a Sheet model because I find that it is not necessary in terms of displaying data. In particular, you are at the mercy of your window size and cannot easily change the dimensions. Also, it complicates the VBA side of things to work in both formats all the time, so I just decided to always put my Charts on Sheets. Your mileage may vary so I'll touch on both approaches in the code samples.)"),(0,i.kt)("h3",{id:"a-quick-overview-of-the-object-model"},"a quick overview of the object model"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"ChartObjects")," -> ",(0,i.kt)("inlineCode",{parentName:"li"},"ChartObject")," - this derives from ",(0,i.kt)("inlineCode",{parentName:"li"},"Shape")," and exists when the Chart is on a Worksheet",(0,i.kt)("ul",{parentName:"li"},(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"Chart"),(0,i.kt)("ul",{parentName:"li"},(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"SeriesCollection")," -> ",(0,i.kt)("inlineCode",{parentName:"li"},"Series")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"Axes")," -> ",(0,i.kt)("inlineCode",{parentName:"li"},"Axis")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"ChartArea")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"PlotArea")))))),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"ActiveChart")," -> ",(0,i.kt)("inlineCode",{parentName:"li"},"Chart")," - this works whether you have a Worksheet or Chart on a sheet"),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"Selection")," -> ",(0,i.kt)("inlineCode",{parentName:"li"},"Variant")," - this one can be useful but is often not of the type that you want.")),(0,i.kt)("h3",{id:"obtaining-a-reference-to-a-chart"},"obtaining a reference to a Chart"),(0,i.kt)("p",null,"When working with ",(0,i.kt)("inlineCode",{parentName:"p"},"Charts"),", the first task is typically to get a reference to an existing chart -- unless you are creating a new chart. To obtain a reference to a chart, there are a handful of ways of doing it depending on what your spreadsheet contains and how it's structured."),(0,i.kt)("p",null,"THe main ways to do it are:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"Use the ",(0,i.kt)("inlineCode",{parentName:"li"},"ActiveChart")," object"),(0,i.kt)("li",{parentName:"ul"},"Use the ",(0,i.kt)("inlineCode",{parentName:"li"},"Selection")," object -- this is highly depending on what is selected"),(0,i.kt)("li",{parentName:"ul"},"Use the ",(0,i.kt)("inlineCode",{parentName:"li"},"ChartObjects")," object",(0,i.kt)("ul",{parentName:"li"},(0,i.kt)("li",{parentName:"ul"},"If you know which chart you want, you can supply an index; this works great if there is only a single chart - ",(0,i.kt)("inlineCode",{parentName:"li"},"ChartObjects(1)")),(0,i.kt)("li",{parentName:"ul"},"If you want to do something to all charts, you can iterate this object"),(0,i.kt)("li",{parentName:"ul"},"If you have named the chart (more on that later) you can supply the name as the index - ",(0,i.kt)("inlineCode",{parentName:"li"},'ChartObjects("SomeChart")')))),(0,i.kt)("li",{parentName:"ul"},"The ",(0,i.kt)("inlineCode",{parentName:"li"},"Workbook.Sheets")," object if your charts are contained in their own sheets",(0,i.kt)("ul",{parentName:"li"},(0,i.kt)("li",{parentName:"ul"},"Same as above, you can access via a numeric index, name, or iterate through all of them")))),(0,i.kt)("h4",{id:"activechart"},(0,i.kt)("inlineCode",{parentName:"h4"},"ActiveChart")),(0,i.kt)("p",null,(0,i.kt)("inlineCode",{parentName:"p"},"ActiveChart")," is similar to the other ",(0,i.kt)("inlineCode",{parentName:"p"},"Active"),' objects in that it does about what you expect. The one difference is that the Chart actually has to be selected or have focus in order to be considered "active". This is similar but also different to something like ',(0,i.kt)("inlineCode",{parentName:"p"},"ActiveWorkbook")," where having the workbook open makes it active."),(0,i.kt)("p",null,"Note that ",(0,i.kt)("inlineCode",{parentName:"p"},"ActiveChart")," will work for a ",(0,i.kt)("inlineCode",{parentName:"p"},"Chart")," that is contained on a Worksheet or also for one that is its own Sheet. If the latter case, then ",(0,i.kt)("inlineCode",{parentName:"p"},"ActiveSheet")," and ",(0,i.kt)("inlineCode",{parentName:"p"},"ActiveChart")," will refer to the same object. Side note: this technicality is why you will not get proper Intellisense when using ",(0,i.kt)("inlineCode",{parentName:"p"},"ActiveSheet")," -- that Sheet could technically be a Chart."),(0,i.kt)("p",null,"The nice thing about ",(0,i.kt)("inlineCode",{parentName:"p"},"ActiveChart")," is that it gives you the Chart object which then gives you immediate access to the Chart related details you are like to want to change. The downside is that unless you have a single Chart that is already selected, ",(0,i.kt)("inlineCode",{parentName:"p"},"ActiveChart")," has limited application when using VBA. Again, the goal is to avoid selecting objects in order to access them via VBA so ",(0,i.kt)("inlineCode",{parentName:"p"},"ActiveChart")," is not ideal."),(0,i.kt)("h4",{id:"selection"},(0,i.kt)("inlineCode",{parentName:"h4"},"Selection")),(0,i.kt)("p",null,"The Selection object is probably the greatest catch all for an object. It literally holds anything, and this means that using the object requires knowing what is selected, or checking vigorously before using the object. Technically, you also let your code error out if the wrong object is selected, and this works well at times. This works well because you are unlikely to be using ",(0,i.kt)("inlineCode",{parentName:"p"},"Selection")," in a complicated workflow because, again, you should not be selecting objects to access them. This means that ",(0,i.kt)("inlineCode",{parentName:"p"},"Selection")," is really limited to one-off and helper code where you can more tightly dictate that this code only works if you select a Chart. You should still add some error handling, but sometimes that step is skipped."),(0,i.kt)("p",null,"Since the ",(0,i.kt)("inlineCode",{parentName:"p"},"Selection")," can hold anything, it's important to know what could be Selected. Related to charts, the following can all live in the ",(0,i.kt)("inlineCode",{parentName:"p"},"Selection"),":"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"ChartObjects")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"Chart")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"ChartArea")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"PlotArea")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"Legend")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"ChartTitle")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"Series"))),(0,i.kt)("p",null,"If you are writing VBA to work on Charts, you can technically require the user to select the correct part of the chart and always use ",(0,i.kt)("inlineCode",{parentName:"p"},"Selection"),". You will quickly grow tired of having to remember which part of the Chart to select in order to make the code work. To avoid this scenario, it is helpful to remember the object model and know how to work your way around a Chart."),(0,i.kt)("p",null,"My approach has always been to convert the ",(0,i.kt)("inlineCode",{parentName:"p"},"Selection")," to a Collection of ",(0,i.kt)("inlineCode",{parentName:"p"},"ChartObjects"),". I can then always iterate that resulting Collection to process the Charts. If only a single Chart was selected, the code works all the same. The downside to this approach is that a Chart as a Sheet cannot live inside a ChartObject. This is a large part of why I always put Charts on a Worksheet."),(0,i.kt)("p",null,"Below is the helper function I use in order to convert a possibly Chart containing selection into a Collection of ",(0,i.kt)("inlineCode",{parentName:"p"},"ChartObjects"),". It works for all objects except for the Axis related ones."),(0,i.kt)("p",null,"TODO: consider improving this code if it is included as a de facto reference"),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-vb"},'Public Function Chart_GetObjectsFromObject(ByVal inputObject As Object) As Variant\n\n    Dim chartObjectCollection As New Collection\n\n    \'NOTE that this function does not work well with Axis objects.  Excel does not return the correct Parent for them.\n\n    Dim targetObject As Variant\n    Dim inputObjectType As String\n    inputObjectType = TypeName(inputObject)\n\n    Select Case inputObjectType\n\n        Case "DrawingObjects"\n            \'this means that multiple charts are selected\n            For Each targetObject In inputObject\n                If TypeName(targetObject) = "ChartObject" Then\n                    \'add it to the set\n                    chartObjectCollection.Add targetObject\n                End If\n            Next targetObject\n\n        Case "Worksheet"\n            For Each targetObject In inputObject.ChartObjects\n                chartObjectCollection.Add targetObject\n            Next targetObject\n\n        Case "Chart"\n            chartObjectCollection.Add inputObject.Parent\n\n        Case "ChartArea", "PlotArea", "Legend", "ChartTitle"\n            \'parent is the chart, parent of that is the chart targetObject\n            chartObjectCollection.Add inputObject.Parent.Parent\n\n        Case "Series"\n            \'need to go up three levels\n            chartObjectCollection.Add inputObject.Parent.Parent.Parent\n\n        Case "Axis", "Gridlines", "AxisTitle"\n            \'these are the oddly unsupported objects\n            MsgBox "Axis/gridline selection not supported.  This is an Excel bug.  Select another element on the chart(s)."\n\n        Case Else\n            MsgBox "Select a part of the chart(s), except an axis."\n\n    End Select\n\n    Set Chart_GetObjectsFromObject = chartObjectCollection\nEnd Function\n')),(0,i.kt)("h4",{id:"chartobjects"},"ChartObjects"),(0,i.kt)("p",null,"If you are working on a Worksheet, then that Worksheet will have the ",(0,i.kt)("inlineCode",{parentName:"p"},"ChartObjects")," object. This object is great because it contains all of the Charts in their own collection (separate from any other Shapes or buttons). This ",(0,i.kt)("inlineCode",{parentName:"p"},"ChartObjects")," collection contains object of type ChartObject. The ChartObject derives from Shape which means it contains all of the properties related to on-sheet position and size."),(0,i.kt)("p",null,"A typical workflow is included below since it is a pattern that shows up all the time in VBA code related to charts. At a high level the steps are:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"Use ActiveSheet or a Worksheet reference to access the ",(0,i.kt)("inlineCode",{parentName:"li"},"ChartObjects")),(0,i.kt)("li",{parentName:"ul"},"Iterate through each ",(0,i.kt)("inlineCode",{parentName:"li"},"ChartObject"),", storing a reference to the underlying Chart"),(0,i.kt)("li",{parentName:"ul"},"You then setup sections to work through the parts of the Chart you want",(0,i.kt)("ul",{parentName:"li"},(0,i.kt)("li",{parentName:"ul"},"Iterate through the ",(0,i.kt)("inlineCode",{parentName:"li"},"SeriesCollection")),(0,i.kt)("li",{parentName:"ul"},"Iterate through the Axes"),(0,i.kt)("li",{parentName:"ul"},"Touch the other top level properties including ",(0,i.kt)("inlineCode",{parentName:"li"},"ChartTile"),", ",(0,i.kt)("inlineCode",{parentName:"li"},"Legend"),", etc.")))),(0,i.kt)("p",null,"This workflow is quite powerful because it can quickly be wrapped with a loop to go through all Worksheets and even possible all Workbooks. It's also powerful because you can be quite comfortable learning this pattern and then adding in the parts that you actually want to change. The only downside is that it can be quite tedious to type out all the loops every time, but there's not a good way around that other than to use the clipboard."),(0,i.kt)("p",null,"Another approach to using ",(0,i.kt)("inlineCode",{parentName:"p"},"ChartObjects")," is to not iterate through all of them but instead to select a single ChartObject and work with it. There are two ways to do this:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"Use an integer index for the Chart -- this is quite easy to do if there are only a few charts"),(0,i.kt)("li",{parentName:"ul"},"Name the chart and use that name")),(0,i.kt)("p",null,"When using either of these approaches, it is quite helpful to show the ",(0,i.kt)("inlineCode",{parentName:"p"},"Selection Pane")," window in Excel. This pane will pop out and tell you the order and the names of all the objects on the sheet (this includes comments, shapes, and Charts). From this pane, you can rearrange the charts into the order you want or rename them."),(0,i.kt)("p",null,"Although ",(0,i.kt)("inlineCode",{parentName:"p"},"For Each")," loops are generally preferred when working with Charts, sometimes you simply know that you want to change one chart and an index just lets you do that. If you are in the habit of using loops however, you can easily do that with the helper code included above which stick a single chart into a Collection."),(0,i.kt)("h4",{id:"workbooksheets-to-get-chart-references"},"Workbook.Sheets to get Chart references"),(0,i.kt)("p",null,"The final approach to obtaining a Chart reference is to use the ",(0,i.kt)("inlineCode",{parentName:"p"},"Sheets")," object. Aside from ActiveChart, this is the only way to deal with Charts that are their own Sheet. Again, you can either use an index or a Name. Here, the Name is easily changed on the Sheet tab so it's much more common to use a Name when doing this. The other approach is to iterate through all the ",(0,i.kt)("inlineCode",{parentName:"p"},"Sheets")," and pick off the ones that are Charts."),(0,i.kt)("p",null,"There are two key points when working with Charts as ",(0,i.kt)("inlineCode",{parentName:"p"},"Sheets"),":"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"You must use the ",(0,i.kt)("inlineCode",{parentName:"li"},"Workbook.Sheets")," object to access them and not ",(0,i.kt)("inlineCode",{parentName:"li"},"Workbook.Worksheets"),". The latter object contains only those ",(0,i.kt)("inlineCode",{parentName:"li"},"Worksheets")," that are not Charts. The former contains both Charts and ",(0,i.kt)("inlineCode",{parentName:"li"},"Worksheets"),"."),(0,i.kt)("li",{parentName:"ul"},"It's possible that your Sheet is not actually a Chart. You should check the type of the object if you are going to iterate through all ",(0,i.kt)("inlineCode",{parentName:"li"},"Worksheets"),". Also be aware that some sheets can be hidden which might lead to unexpected results.")),(0,i.kt)("p",null,"TODO: is there a Charts object on Workbook?"))}u.isMDXComponent=!0}}]);