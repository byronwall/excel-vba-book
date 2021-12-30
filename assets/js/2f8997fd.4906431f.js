"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[4471],{3905:function(e,t,n){n.d(t,{Zo:function(){return c},kt:function(){return p}});var o=n(7294);function a(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function i(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);t&&(o=o.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,o)}return n}function r(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?i(Object(n),!0).forEach((function(t){a(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):i(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function s(e,t){if(null==e)return{};var n,o,a=function(e,t){if(null==e)return{};var n,o,a={},i=Object.keys(e);for(o=0;o<i.length;o++)n=i[o],t.indexOf(n)>=0||(a[n]=e[n]);return a}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(o=0;o<i.length;o++)n=i[o],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(a[n]=e[n])}return a}var l=o.createContext({}),d=function(e){var t=o.useContext(l),n=t;return e&&(n="function"==typeof e?e(t):r(r({},t),e)),n},c=function(e){var t=d(e.components);return o.createElement(l.Provider,{value:t},e.children)},u={inlineCode:"code",wrapper:function(e){var t=e.children;return o.createElement(o.Fragment,{},t)}},h=o.forwardRef((function(e,t){var n=e.components,a=e.mdxType,i=e.originalType,l=e.parentName,c=s(e,["components","mdxType","originalType","parentName"]),h=d(n),p=a,f=h["".concat(l,".").concat(p)]||h[p]||u[p]||i;return n?o.createElement(f,r(r({ref:t},c),{},{components:n})):o.createElement(f,r({ref:t},c))}));function p(e,t){var n=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var i=n.length,r=new Array(i);r[0]=h;var s={};for(var l in t)hasOwnProperty.call(t,l)&&(s[l]=t[l]);s.originalType=e,s.mdxType="string"==typeof e?e:a,r[1]=s;for(var d=2;d<i;d++)r[d]=n[d];return o.createElement.apply(null,r)}return o.createElement.apply(null,n)}h.displayName="MDXCreateElement"},438:function(e,t,n){n.r(t),n.d(t,{frontMatter:function(){return s},contentTitle:function(){return l},metadata:function(){return d},toc:function(){return c},default:function(){return h}});var o=n(7462),a=n(3366),i=(n(7294),n(3905)),r=["components"],s={},l=void 0,d={unversionedId:"overview-of-building-an-addin/14-03 specific-aspects-to-addin-development",id:"overview-of-building-an-addin/14-03 specific-aspects-to-addin-development",title:"14-03 specific-aspects-to-addin-development",description:"specific aspects to addin development",source:"@site/docs/14-overview-of-building-an-addin/14-03 specific-aspects-to-addin-development.md",sourceDirName:"14-overview-of-building-an-addin",slug:"/overview-of-building-an-addin/14-03 specific-aspects-to-addin-development",permalink:"/excel-vba-book/docs/overview-of-building-an-addin/14-03 specific-aspects-to-addin-development",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/14-overview-of-building-an-addin/14-03 specific-aspects-to-addin-development.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"14-02 creating-an-addin",permalink:"/excel-vba-book/docs/overview-of-building-an-addin/14-02 creating-an-addin"},next:{title:"14-04 UI-features-for-addins-Ribbon-toolbars-UserForms",permalink:"/excel-vba-book/docs/overview-of-building-an-addin/14-04 UI-features-for-addins-Ribbon-toolbars-UserForms"}},c=[{value:"specific aspects to addin development",id:"specific-aspects-to-addin-development",children:[{value:"Keyboard Shortcuts",id:"keyboard-shortcuts",children:[],level:3},{value:"USer Forms",id:"user-forms",children:[],level:3},{value:"Helpful Commands",id:"helpful-commands",children:[],level:3},{value:"Other functionality",id:"other-functionality",children:[],level:3}],level:2}],u={toc:c};function h(e){var t=e.components,n=(0,a.Z)(e,r);return(0,i.kt)("wrapper",(0,o.Z)({},u,n,{components:t,mdxType:"MDXLayout"}),(0,i.kt)("h2",{id:"specific-aspects-to-addin-development"},"specific aspects to addin development"),(0,i.kt)("p",null,"Depending on the addin that you are creating, you may expect for it to have a handful of features available. In general, those types of features include keyboard shortcuts, special forms or user prompts, and possibly automatic features that fire depending on the user's action or the state of the workbook or Application."),(0,i.kt)("h3",{id:"keyboard-shortcuts"},"Keyboard Shortcuts"),(0,i.kt)("p",null,"The simplest thing to do is to add keyboard shortcuts to your addin. There are two ways to do that:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},'Open up the Macros form on the Developer tab. You can then hit "options" for a given Sub and assign a keyboard shortcut (TODO: add picture of this)'),(0,i.kt)("li",{parentName:"ul"},"That approach can sometimes be a pain to edit later, so you can also add code to your addin to add the shortcut.")),(0,i.kt)("p",null,"The latter approach is nice because you can easily change the shortcut or the calling method. For addins, I will nearly always take the latter approach since it is much easier to deal with alter. For XLSM workbooks, I will do the former since it is easier to change from a workbook."),(0,i.kt)("p",null,"If you want to add the keyboard shortcut using code, use the code below. Ideally, you would put this in a Workbook_Open event that is called when the workbook opens. You can also use this approach to add/remove shortcuts depending on user input."),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-vb"},'Public Sub SetUpKeyboardHooksForSelection()\n\n\n    \'SHIFT =    +\n    \'CTRL =     ^\n    \'ALT =      %\n\n    \'set up the keys for the selection mover\n    Application.OnKey "^%{RIGHT}", "SelectionOffsetRight"\n    Application.OnKey "^%{LEFT}", "SelectionOffsetLeft"\n    Application.OnKey "^%{UP}", "SelectionOffsetUp"\n    Application.OnKey "^%{DOWN}", "SelectionOffsetDown"\n\n    \'set up the keys for the indent level\n    Application.OnKey "+^%{RIGHT}", "Formatting_IncreaseIndentLevel"\n    Application.OnKey "+^%{LEFT}", "Formatting_DecreaseIndentLevel"\n\nEnd Sub\n')),(0,i.kt)("h3",{id:"user-forms"},"USer Forms"),(0,i.kt)("p",null,"One of the nice features of an addin are adding custom forms to provide the user with a better experience. Creating a UserForm in VBA is dead simple, and this is the best bang for your buck in terms of creating a professional looking product. The simplest of forms with the simplest of features can save the end user hours and hours of time (I've seen it happen)."),(0,i.kt)("p",null,"The nice thing here is that creating a UserForm in an addin is not any different than creating them normally. You simply create the UserForm. The only extra step is that you need to manage how/when the form is created and what information it has access to. Typically this is done by adding a button or using a keyboard shortcut. The only other issue is that you need to be aware of which Workbook or Worksheet is active when opening a UserForm if you are using ActiveSheet or ActiveWorkbook for anything. In general, inside an addin, you need to be careful with this commands since it is not always obvious that the ActiveXXX is the one you want to access."),(0,i.kt)("h3",{id:"helpful-commands"},"Helpful Commands"),(0,i.kt)("p",null,"There are a couple of commands that exist outside of addins that become far more useful inside the addin. They are included below for reference:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"ThisWorkbook")," refers to the workbook that contains the code being executed. This is the surefire way to refer to the XLAM file that is running instead of the ActiveWorkbook. IN general, your addin will never be the ActiveWorkbook. This becomes relevant if your addin workbook contains sheets of data that may need to be accessed during runtime. You would use THisWorkbook to refer to those sheet."),(0,i.kt)("li",{parentName:"ul"},"TODO: add any other commands that are addin specific")),(0,i.kt)("h3",{id:"other-functionality"},"Other functionality"),(0,i.kt)("p",null,"THe other functionality that you can add is related to Events. You have great power when it comes to listening to events and triggering various actions. THe real difficulty is deciding what is an appropriate use of that power. Namely, when will you create an experience that benefits the user versus creating a very confusing workbook that is prone to breaking?"),(0,i.kt)("p",null,"Before diving into what events can do, it's worth nting that potential downfalls of using them:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"They can be quite finicky sometimes. That is, using events adds a layer of complexity that tends to just complicate Excel and VBA. I don't have a technical explanation, but there seem to be a number of bugs that creep out of the dark once you start really using events."),(0,i.kt)("li",{parentName:"ul"},"Your user can disable events at will and it can be quite difficult to determine when that was done. This is done with ",(0,i.kt)("inlineCode",{parentName:"li"},"Application.EnableEvents = False"),"."),(0,i.kt)("li",{parentName:"ul"},"Events are triggered all the time for all sorts of reasons. If you are doing a lot of checking in Events, you will dramatically slow down the workbook.")),(0,i.kt)("p",null,"With all of those warnings, there is nothing wrong with using Events. They generally do what you want and can be quite powerful. I add the caveats only because I have seen them ruin an otherwise working workbook. That complexity gets amped up a level when your Event code is inside an addin instead of the main workbook."),(0,i.kt)("p",null,'To really make the most of Events, you are going to need to use Class Modules. The reason is that your Events need to "latch on" to the host workbooks or worksheets, and the only way to do that is by using Class Modules. Normally, outside of an addin, you can simply open up the relevant VBA object (Workbook or Worksheet) and add the event code there. For an addin, you cannot add that code outside of the addin so you are in a bind. How then can you hook onto the Event? Fortunately, VBA makes this possible with the ',(0,i.kt)("inlineCode",{parentName:"p"},"With Events")," command inside of a Class Module."),(0,i.kt)("p",null,"TODO: provide a concrete example of using this code"))}h.isMDXComponent=!0}}]);