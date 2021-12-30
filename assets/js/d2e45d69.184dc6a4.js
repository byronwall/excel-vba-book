"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[3529],{3905:function(e,t,a){a.d(t,{Zo:function(){return u},kt:function(){return p}});var n=a(7294);function r(e,t,a){return t in e?Object.defineProperty(e,t,{value:a,enumerable:!0,configurable:!0,writable:!0}):e[t]=a,e}function o(e,t){var a=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);t&&(n=n.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),a.push.apply(a,n)}return a}function i(e){for(var t=1;t<arguments.length;t++){var a=null!=arguments[t]?arguments[t]:{};t%2?o(Object(a),!0).forEach((function(t){r(e,t,a[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(a)):o(Object(a)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(a,t))}))}return e}function s(e,t){if(null==e)return{};var a,n,r=function(e,t){if(null==e)return{};var a,n,r={},o=Object.keys(e);for(n=0;n<o.length;n++)a=o[n],t.indexOf(a)>=0||(r[a]=e[a]);return r}(e,t);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);for(n=0;n<o.length;n++)a=o[n],t.indexOf(a)>=0||Object.prototype.propertyIsEnumerable.call(e,a)&&(r[a]=e[a])}return r}var l=n.createContext({}),c=function(e){var t=n.useContext(l),a=t;return e&&(a="function"==typeof e?e(t):i(i({},t),e)),a},u=function(e){var t=c(e.components);return n.createElement(l.Provider,{value:t},e.children)},h={inlineCode:"code",wrapper:function(e){var t=e.children;return n.createElement(n.Fragment,{},t)}},d=n.forwardRef((function(e,t){var a=e.components,r=e.mdxType,o=e.originalType,l=e.parentName,u=s(e,["components","mdxType","originalType","parentName"]),d=c(a),p=r,f=d["".concat(l,".").concat(p)]||d[p]||h[p]||o;return a?n.createElement(f,i(i({ref:t},u),{},{components:a})):n.createElement(f,i({ref:t},u))}));function p(e,t){var a=arguments,r=t&&t.mdxType;if("string"==typeof e||r){var o=a.length,i=new Array(o);i[0]=d;var s={};for(var l in t)hasOwnProperty.call(t,l)&&(s[l]=t[l]);s.originalType=e,s.mdxType="string"==typeof e?e:r,i[1]=s;for(var c=2;c<o;c++)i[c]=a[c];return n.createElement.apply(null,i)}return n.createElement.apply(null,a)}d.displayName="MDXCreateElement"},2184:function(e,t,a){a.r(t),a.d(t,{frontMatter:function(){return s},contentTitle:function(){return l},metadata:function(){return c},toc:function(){return u},default:function(){return d}});var n=a(7462),r=a(3366),o=(a(7294),a(3905)),i=["components"],s={},l=void 0,c={unversionedId:"overview-of-basics-of-VBA/03-03 declaring-and-setting-variables",id:"overview-of-basics-of-VBA/03-03 declaring-and-setting-variables",title:"03-03 declaring-and-setting-variables",description:"Declaring and Setting Variables",source:"@site/docs/03-overview-of-basics-of-VBA/03-03 declaring-and-setting-variables.md",sourceDirName:"03-overview-of-basics-of-VBA",slug:"/overview-of-basics-of-VBA/03-03 declaring-and-setting-variables",permalink:"/docs/overview-of-basics-of-VBA/03-03 declaring-and-setting-variables",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/03-overview-of-basics-of-VBA/03-03 declaring-and-setting-variables.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"03-02 introduction-to-VBA",permalink:"/docs/overview-of-basics-of-VBA/03-02 introduction-to-VBA"},next:{title:"03-04 using-Subs-and-Functions",permalink:"/docs/overview-of-basics-of-VBA/03-04 using-Subs-and-Functions"}},u=[{value:"Declaring and Setting Variables",id:"declaring-and-setting-variables",children:[{value:"Declaring Variables",id:"declaring-variables",children:[],level:3},{value:"Setting variables",id:"setting-variables",children:[],level:3},{value:"Using Variables",id:"using-variables",children:[],level:3},{value:"Value Default",id:"value-default",children:[],level:3}],level:2}],h={toc:u};function d(e){var t=e.components,a=(0,r.Z)(e,i);return(0,o.kt)("wrapper",(0,n.Z)({},h,a,{components:t,mdxType:"MDXLayout"}),(0,o.kt)("h2",{id:"declaring-and-setting-variables"},"Declaring and Setting Variables"),(0,o.kt)("p",null,"One of the core tasks when programming via VBA is working with variables. Variables are used to reference the Excel object model and to guide control structures. Within the Excel object model, the objects hold variables which point to other objects. Working with these objects is critical to using VBA. You will need to understand variables to do that."),(0,o.kt)("p",null,"This section is split into two areas: declaring variables and setting variables. The code for these two topics is simple. The complexity comes in planning out the best structure for managing variables. The variable declaration will directly shape how the control structures will work."),(0,o.kt)("h3",{id:"declaring-variables"},"Declaring Variables"),(0,o.kt)("p",null,"Declaring variables is straight forward. VBA offers a simple command to declare a new variable: ",(0,o.kt)("inlineCode",{parentName:"p"},"Dim"),"."),(0,o.kt)("p",null,"When declaring a variable, there are two components to it: variable name and variable type. Variable names are your choice with some constraints. You are not allowed to duplicate the name of an internal command, and you should go to some length to avoid using the same name as an Excel object model name. Beware that naming a variable has certain conventions, but these do not have any effect on the program execution. The main concern with names is that they will directly affect your ability to work with and maintain your code. Naming things is hard. Pick a strategy that works for you and your coworkers and get on it with it. There is no single answer here about how to name things."),(0,o.kt)("p",null,"The second part of the puzzle is to declare the type of the variable. This is THE core part of variables. When declaring a variable, you decide if the type should be the generic ",(0,o.kt)("inlineCode",{parentName:"p"},"Variant")," or if you need a more specific type. There are times when you have to use Variant, but you should aim to use the most specific type that is possible. These types draw from VBA, from the Excel Object Model, or from your own created types. When thinking of variable types, there are two major groups of types:"),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"Value types = a number, string, or boolean"),(0,o.kt)("li",{parentName:"ul"},"Reference types = objects")),(0,o.kt)("p",null,"TODO: find better place:"),(0,o.kt)("p",null,"Note that you can technically use a variable before declaring it, but you should really avoid this practice. It leads to the potential to create all sorts of bugs later. Just don't do it. To better avoid this, setting the flag in the settings (TODO: add a picture of that)."),(0,o.kt)("p",null,"TODO: add code sample for declaring a variable (show an object, primitive, and array)"),(0,o.kt)("h3",{id:"setting-variables"},"Setting variables"),(0,o.kt)("p",null,"Setting a variable is straight forward. The rule is: ",(0,o.kt)("strong",{parentName:"p"},"for reference types, you must use ",(0,o.kt)("inlineCode",{parentName:"strong"},"Set"),"; for value types, you must not.")),(0,o.kt)("p",null,"The real problem then is to determine whether or not you are working with a reference type. The rule is: if you are working with an object, it is a reference type. If you are working with a value (number, string, boolean), then you have a value type. Another approach, if you intend to use a ",(0,o.kt)("inlineCode",{parentName:"p"},".")," to call out some property of your variable, then it is a reference type. The exception here is arrays: they are set without using Set."),(0,o.kt)("p",null,"TODO: add code sample showing variable setting"),(0,o.kt)("h3",{id:"using-variables"},"Using Variables"),(0,o.kt)("p",null,"It seems somewhat obvious that you would want to use a variable after declaring and setting it. This is generally always the case (why else would you create the variable). To that end, there are a pair of ways to use variables depending on whether it is a reference or value type. Value types are easier since you can only do 1 thing with them: use them in an expression. This feels and usually looks like mathematical formulas. The more complicated example comes with reference types where the variable stores a reference to another object. These variables have the ability to access either a property of the type or the default ",(0,o.kt)("inlineCode",{parentName:"p"},"Value")," of the type. The distinctions between reference and value types can become confusing with the Excel Object Model since so many properties of objects reduce to value types. An example is the value of a ",(0,o.kt)("inlineCode",{parentName:"p"},"Range")," which will hold some number or string or Error depending on what the cell contains."),(0,o.kt)("p",null,"When accessing a property of the object, you use the ",(0,o.kt)("inlineCode",{parentName:"p"},".")," to access a property by name. In this way, you can chain together a series of commands accessing the properties of objects. It is often the case that the property is itself another object which makes it possible to use another ",(0,o.kt)("inlineCode",{parentName:"p"},".")," to keep going. If you are using the VBE and properly declaring your variables, the VBE will work to provide helpful suggestions of what may be possible to use next (this is called Intellisense). The one pitfall to Intellisense is when the return from a given property can be Variant or a combination of possible results. When this happens, Intellisense will not offer any suggestions and you are left guessing whether or not the command exists. This is where it can be quite helpful to do one of two things:"),(0,o.kt)("p",null,"TODO: create a demo of these bullets"),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},'Create a new variable with the type that you know the object will have and Set that reference before using it. This "cheats" and tells Intellisense exactly what you expect to exist.'),(0,o.kt)("li",{parentName:"ul"},"Read through the documentation and gain an understanding of what types are possible and just use them. There is no rule that the type must be suggested by Intellisense for it to be valid.")),(0,o.kt)("p",null,"In general, I take a combination of those two approaches often. If I expect to use the variable a number of times, I will go with the new variable route to avoid guessing properties later. If I only need the variable once or am copying code from somewhere else (and know it works), I will just go with the code as is without Intellisense. The one upside of creating new variables is that it forces you to be more explicit with your declarations. It also clearly shows your intent to other developers that may see your code later."),(0,o.kt)("h3",{id:"value-default"},"Value Default"),(0,o.kt)("p",null,"I mentioned it above, but it is worth digging into the default ",(0,o.kt)("inlineCode",{parentName:"p"},"Value")," property a little more. This can be a source of confusion because very often, you will accidentally use the name of a variable without calling for a property. In other programming languages, this will result in a compile time or runtime error. In VBA, your code will run and even worse will return something from the object that may not be what you want. When this happens, it can be incredibly difficult to track down the source of the error. To avoid this, you could never use the variable name as a shortcut to the ",(0,o.kt)("inlineCode",{parentName:"p"},".Value")," property. In practice this is a pain to manage and I will often mix and match whether or not Value is called. Sometimes, I am tired of typing out Value and just let the default work. Other times, I am being very diligent about calling everything explicitly to avoid some unforeseen error later. You will find that this comes down to your own preference and the preferences of others working on your code."))}d.isMDXComponent=!0}}]);