"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[4155],{3905:function(e,t,o){o.d(t,{Zo:function(){return u},kt:function(){return d}});var n=o(7294);function a(e,t,o){return t in e?Object.defineProperty(e,t,{value:o,enumerable:!0,configurable:!0,writable:!0}):e[t]=o,e}function l(e,t){var o=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);t&&(n=n.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),o.push.apply(o,n)}return o}function i(e){for(var t=1;t<arguments.length;t++){var o=null!=arguments[t]?arguments[t]:{};t%2?l(Object(o),!0).forEach((function(t){a(e,t,o[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(o)):l(Object(o)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(o,t))}))}return e}function r(e,t){if(null==e)return{};var o,n,a=function(e,t){if(null==e)return{};var o,n,a={},l=Object.keys(e);for(n=0;n<l.length;n++)o=l[n],t.indexOf(o)>=0||(a[o]=e[o]);return a}(e,t);if(Object.getOwnPropertySymbols){var l=Object.getOwnPropertySymbols(e);for(n=0;n<l.length;n++)o=l[n],t.indexOf(o)>=0||Object.prototype.propertyIsEnumerable.call(e,o)&&(a[o]=e[o])}return a}var s=n.createContext({}),c=function(e){var t=n.useContext(s),o=t;return e&&(o="function"==typeof e?e(t):i(i({},t),e)),o},u=function(e){var t=c(e.components);return n.createElement(s.Provider,{value:t},e.children)},p={inlineCode:"code",wrapper:function(e){var t=e.children;return n.createElement(n.Fragment,{},t)}},h=n.forwardRef((function(e,t){var o=e.components,a=e.mdxType,l=e.originalType,s=e.parentName,u=r(e,["components","mdxType","originalType","parentName"]),h=c(o),d=a,m=h["".concat(s,".").concat(d)]||h[d]||p[d]||l;return o?n.createElement(m,i(i({ref:t},u),{},{components:o})):n.createElement(m,i({ref:t},u))}));function d(e,t){var o=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var l=o.length,i=new Array(l);i[0]=h;var r={};for(var s in t)hasOwnProperty.call(t,s)&&(r[s]=t[s]);r.originalType=e,r.mdxType="string"==typeof e?e:a,i[1]=r;for(var c=2;c<l;c++)i[c]=o[c];return n.createElement.apply(null,i)}return n.createElement.apply(null,o)}h.displayName="MDXCreateElement"},7039:function(e,t,o){o.r(t),o.d(t,{frontMatter:function(){return r},contentTitle:function(){return s},metadata:function(){return c},toc:function(){return u},default:function(){return h}});var n=o(7462),a=o(3366),l=(o(7294),o(3905)),i=["components"],r={},s=void 0,c={unversionedId:"The-Application-object/09-02 Controlling-calculations",id:"The-Application-object/09-02 Controlling-calculations",title:"09-02 Controlling-calculations",description:"Controlling calculations",source:"@site/docs/09-The-Application-object/09-02 Controlling-calculations.md",sourceDirName:"09-The-Application-object",slug:"/The-Application-object/09-02 Controlling-calculations",permalink:"/docs/The-Application-object/09-02 Controlling-calculations",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/09-The-Application-object/09-02 Controlling-calculations.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"09-01 introduction-to-the-Application",permalink:"/docs/The-Application-object/09-01 introduction-to-the-Application"},next:{title:"09-03 Controlling-events-and-visuals",permalink:"/docs/The-Application-object/09-03 Controlling-events-and-visuals"}},u=[{value:"Controlling calculations",id:"controlling-calculations",children:[{value:"Disabling calculations",id:"disabling-calculations",children:[],level:3}],level:2}],p={toc:u};function h(e){var t=e.components,o=(0,a.Z)(e,i);return(0,l.kt)("wrapper",(0,n.Z)({},p,o,{components:t,mdxType:"MDXLayout"}),(0,l.kt)("h2",{id:"controlling-calculations"},"Controlling calculations"),(0,l.kt)("p",null,"When you are creating macro workflows, there are a number of tools at your disposal to control calculations flow. Before describing those tools, it's worth stepping back and discussing why you might want to control the calculation flow. There are a couple of common reason:"),(0,l.kt)("ul",null,(0,l.kt)("li",{parentName:"ul"},"Performance. Your code will run faster if you control the calculation process. This mainly involves disabling automatic calculation at key points."),(0,l.kt)("li",{parentName:"ul"},"Accuracy. For some types of calculations, you need to tightly control the calculation flow for accuracy. This is often the case if you are building a spreadsheet that does some form of recursion or self reference."),(0,l.kt)("li",{parentName:"ul"},"Usability. There are some situations where you are interacting with calculations and need to prevent the normal behavior. The most common is when you add Workbook events like ",(0,l.kt)("inlineCode",{parentName:"li"},"Change"),"."),(0,l.kt)("li",{parentName:"ul"},"Profiling. If you are building a code profiler (i.e. a tool that tracks execution time of your code) you must control calculations in order to get the tracking right.")),(0,l.kt)("p",null,"We'll get back to the applications, but it's also worth hitting the high points on how you can control the calculation. THe main knobs:"),(0,l.kt)("ul",null,(0,l.kt)("li",{parentName:"ul"},"Disable application wide"),(0,l.kt)("li",{parentName:"ul"},"Disable for a Worksheet"),(0,l.kt)("li",{parentName:"ul"},"Manually calculate a Range, Worksheet, or Application")),(0,l.kt)("p",null,"THe types of changes you will make are fairly tightly couple to the applications above. In general, for performances nad usability reasons, you will be disable calculations. For accuracy or profiling applications, you will manually walking the calculation through."),(0,l.kt)("h3",{id:"disabling-calculations"},"Disabling calculations"),(0,l.kt)("p",null,'The most common approach to controlling calculations is to simply disable them. To "disable" the calculations, is really to set the CalculationMode to Manual. It does not actually disable calculations, but instead it prevents the automatic calculations updates from firing like normal. The spreadsheet still maintains its normal model of calculations; they just don\'t run. This is an incredibly common approach to speeding up the performance of VBA code. The performance boost results form the fact that when VBA code executes, it is very tightly coupled to the normal Excel operations that take place. When you use VBA to set a ',(0,l.kt)("inlineCode",{parentName:"p"},".Value")," equal to some new value, it is functionally equivlabet to manually entering the value. Behind the scenes, Excel will fire off the normal Change events and update the dependent cells. This can become a bottleneck because VBA is able to rapidly fire off ",(0,l.kt)("inlineCode",{parentName:"p"},".Value")," changes. So rapidly, that processing all of the associated stuff can become a limitation. It is more or less guaranteed that you will run into this issue once you start writing VBA code. It is so common, that you will likely memorize the fix:"),(0,l.kt)("p",null,"TODO: check this code"),(0,l.kt)("pre",null,(0,l.kt)("code",{parentName:"pre",className:"language-vba"},"Application.CalculationMode = xlManual\nApplication.ScreenUpdating = False\nApplication.EnabledEvents = False\n\nApplication.EnabledEvents = True\nApplication.ScreenUpdating = True\nApplication.CalculationMode = xlAutomatic\n")),(0,l.kt)("p",null,"Why does this code make everything faster? Well, it disables the slowest steps of Excel keeping track of your spreadsheet: visual updates, calculations updates, and other events. Turning all of those off will dramatically remove the bottlenecks to your code. What's the downside? Well, all of that stuff exists for a reason and it's possible you need it to keep functioning for some VBA operations. The non-calculatojn options ar ecovered in subsequent chapters, so we'll focus on the calculation part now."),(0,l.kt)("p",null,"What happens when you disable calculations? This is the key concept to understand to make sure your spreadsheets do not break when you go looking for performances. So what changes?"),(0,l.kt)("ul",null,(0,l.kt)("li",{parentName:"ul"},'Dependent cells are not updated. The "chain" is processed to its end on every update. Note, updates are sent downstream. Not all cells are updated, unless your Workbook contains a VOLATILE function.'),(0,l.kt)("li",{parentName:"ul"},"Charts and other functional graphics do not update. Internally, they don't change at all. It's not just a matter of the visuals being hidden, they are not calculated."),(0,l.kt)("li",{parentName:"ul"},"Less important items:",(0,l.kt)("ul",{parentName:"li"},(0,l.kt)("li",{parentName:"ul"},"Conditional formatting will not update.")))),(0,l.kt)("p",null,"So why might those things matter? The biggest reason is that if your VBA code depends on the state of the spreadsheets, then you are likely depending on calculations at some point. This means that you need to split you rcode into segments where you are not worried about cell values and those where you are. An example:"),(0,l.kt)("p",null,"You are building a tool to process data from a CSV file. You have been told that you should delete data that is in the 0th to 10th percentile of a cost column. Unforateunyl, the data needs to be preprocessed in order to create an accurate cost column. Your CSV file contains a mess of extra text and other issues when need to be removed. Your workflow then is:"),(0,l.kt)("ol",null,(0,l.kt)("li",{parentName:"ol"},"Import the CSV data"),(0,l.kt)("li",{parentName:"ol"},"Preprocess the cost column to clean up the mess"),(0,l.kt)("li",{parentName:"ol"},"Remove the rows below the 10th percentile.")),(0,l.kt)("p",null,"You do a quick test and have no problem importing the CSV data. You've gone ahead and worked out the preprocessing logic... only took a couple calls to ",(0,l.kt)("inlineCode",{parentName:"p"},"Split")," and ",(0,l.kt)("inlineCode",{parentName:"p"},"Trim"),". You also went ahead and added a new column to compute the ",(0,l.kt)("inlineCode",{parentName:"p"},"PERCENTILE"),' based on the now cleaned result. This is looking great on your 100 row test data set. Your set your application loose on the 90,000 "real" data and quickly find that it will not complete within 10 minutes. What\'s going on here? THe most likely problem is that your new PERCENTILE column is being reculated every time a preprocessed data cell is being added back to the spreadsheet. Your processing code looks like:'),(0,l.kt)("pre",null,(0,l.kt)("code",{parentName:"pre",className:"language-vba"},"For Each rngCell in rngData\n    rngCell.Value = CleanUpThisMess(rngCell)\nNext\n")),(0,l.kt)("p",null,"If ",(0,l.kt)("inlineCode",{parentName:"p"},"rngData")," contains 90,000 cells, then your update code will call for at least 90,000 full Worksheet recalculations. Even worse, your PERCENTILE formula requires the entire column of data and so all cells have to update every time. ",(0,l.kt)("inlineCode",{parentName:"p"},"90,000 x 90,000")," quickly becomes a problem."),(0,l.kt)("p",null,"So, why is the PERCENTILE function updating after every change? Do we really care what the intermediate values are? No."),(0,l.kt)("p",null,"This is why you want to have control of the calculations. in this case, you know that the processing code is not affected by the value of the PERCENTILE column. We only need the static data available in order to complete the processing. The fix here is to turn calcuaitliosn to manual during the processing step so that you do not incur 90,000 extra recalculations."),(0,l.kt)("p",null,"Once the processing is done, what do we do with the calculation mode? Well, that depends on how we do the deletion. There are a couple of options:"),(0,l.kt)("ul",null,(0,l.kt)("li",{parentName:"ul"},"Turn on an ",(0,l.kt)("inlineCode",{parentName:"li"},"AutoFilter")," and do a FILTER-DELETE to remove all the rows in one shot."),(0,l.kt)("li",{parentName:"ul"},"Iterate through the rows, one by one, and remove those which are in the 10th or lower percentiles")),(0,l.kt)("p",null,"Looks like either will work, but hwo does calculation mode affect things? Well, if you go with the latter option, you will find that your PERCENTILES will update after each deletion. This is not the behavior you intended. You somehow want to remember the PERCENTILE value before you started the deletions. The solution then is to control the calculation mode again. Here, we are controlling things for ",(0,l.kt)("strong",{parentName:"p"},"accuracy"),". Our deletion approach will not work if we allow cells to update as we go."),(0,l.kt)("p",null,"Pro tip: if you are deleting cells, you should pretty much never go a row, column, or cell at a time. Instead you should build a ",(0,l.kt)("inlineCode",{parentName:"p"},"Range")," of cells to be deleted using ",(0,l.kt)("inlineCode",{parentName:"p"},"Union")," and delete them in one shot using ",(0,l.kt)("inlineCode",{parentName:"p"},"Delete"),". This approach is called a ",(0,l.kt)("inlineCode",{parentName:"p"},"UNION-DELETE")," and avoids all of the issues described above. It's also the fatest approach since it does a single deletion."))}h.isMDXComponent=!0}}]);