"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[6303],{3905:function(e,t,n){n.d(t,{Zo:function(){return c},kt:function(){return h}});var r=n(7294);function a(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function o(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);t&&(r=r.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,r)}return n}function i(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?o(Object(n),!0).forEach((function(t){a(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):o(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function s(e,t){if(null==e)return{};var n,r,a=function(e,t){if(null==e)return{};var n,r,a={},o=Object.keys(e);for(r=0;r<o.length;r++)n=o[r],t.indexOf(n)>=0||(a[n]=e[n]);return a}(e,t);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);for(r=0;r<o.length;r++)n=o[r],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(a[n]=e[n])}return a}var u=r.createContext({}),l=function(e){var t=r.useContext(u),n=t;return e&&(n="function"==typeof e?e(t):i(i({},t),e)),n},c=function(e){var t=l(e.components);return r.createElement(u.Provider,{value:t},e.children)},p={inlineCode:"code",wrapper:function(e){var t=e.children;return r.createElement(r.Fragment,{},t)}},d=r.forwardRef((function(e,t){var n=e.components,a=e.mdxType,o=e.originalType,u=e.parentName,c=s(e,["components","mdxType","originalType","parentName"]),d=l(n),h=a,m=d["".concat(u,".").concat(h)]||d[h]||p[h]||o;return n?r.createElement(m,i(i({ref:t},c),{},{components:n})):r.createElement(m,i({ref:t},c))}));function h(e,t){var n=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var o=n.length,i=new Array(o);i[0]=d;var s={};for(var u in t)hasOwnProperty.call(t,u)&&(s[u]=t[u]);s.originalType=e,s.mdxType="string"==typeof e?e:a,i[1]=s;for(var l=2;l<o;l++)i[l]=n[l];return r.createElement.apply(null,i)}return r.createElement.apply(null,n)}d.displayName="MDXCreateElement"},6878:function(e,t,n){n.r(t),n.d(t,{frontMatter:function(){return s},contentTitle:function(){return u},metadata:function(){return l},toc:function(){return c},default:function(){return d}});var r=n(7462),a=n(3366),o=(n(7294),n(3905)),i=["components"],s={},u=void 0,l={unversionedId:"overview-of-basics-of-VBA/03-05 declaring-the-parameters-Subs-and-Functions",id:"overview-of-basics-of-VBA/03-05 declaring-the-parameters-Subs-and-Functions",title:"03-05 declaring-the-parameters-Subs-and-Functions",description:"declaring the parameters (Subs and Functions)",source:"@site/docs/03-overview-of-basics-of-VBA/03-05 declaring-the-parameters-Subs-and-Functions.md",sourceDirName:"03-overview-of-basics-of-VBA",slug:"/overview-of-basics-of-VBA/03-05 declaring-the-parameters-Subs-and-Functions",permalink:"/excel-vba-book/docs/overview-of-basics-of-VBA/03-05 declaring-the-parameters-Subs-and-Functions",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/03-overview-of-basics-of-VBA/03-05 declaring-the-parameters-Subs-and-Functions.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"03-04 using-Subs-and-Functions",permalink:"/excel-vba-book/docs/overview-of-basics-of-VBA/03-04 using-Subs-and-Functions"},next:{title:"03-06 calling-a-Sub-or-Function",permalink:"/excel-vba-book/docs/overview-of-basics-of-VBA/03-06 calling-a-Sub-or-Function"}},c=[{value:"declaring the parameters (Subs and Functions)",id:"declaring-the-parameters-subs-and-functions",children:[{value:"declaring an Optional parameter",id:"declaring-an-optional-parameter",children:[],level:3}],level:2}],p={toc:c};function d(e){var t=e.components,n=(0,a.Z)(e,i);return(0,o.kt)("wrapper",(0,r.Z)({},p,n,{components:t,mdxType:"MDXLayout"}),(0,o.kt)("h2",{id:"declaring-the-parameters-subs-and-functions"},"declaring the parameters (Subs and Functions)"),(0,o.kt)("p",null,"When creating a new Sub or Function you are able to determine the inputs to your new creation. There are a handful of ways of handling the inputs:"),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"Put the inputs into the parameters of the Sub/Function and allow the caller to provide them"),(0,o.kt)("li",{parentName:"ul"},"Use knowledge of the spreadsheet to determine the inputs (or prompt the user for an input)")),(0,o.kt)("p",null,"The main split here is: do you require the person typing the VBA to give you the inputs? Or, do you use some other approach like asking the user or just pulling the inputs from the spreadsheet."),(0,o.kt)("p",null,"The most common approach is to pull the inputs out of the spreadsheet. This seems counter intuitive, but if you consider that the vast majority of VBA code is purposes written for a single use, then it stands to reason that code will not be built on a large number of Subs/Functions accepting parameters. The reason for this is that generally someone writes VBA to handle ",(0,o.kt)("em",{parentName:"p"},"their")," spreadsheet and so the VBA just reflects that spreadsheet. This works great for individual cases but can become a burden when building larger workflows. The main thing to consider for lager workflows is that as the complexity grows, there will be a large amount og code that is called multiple times or could be called separately from the main workflow. When the sis the case, you are often served by pulling that code out into its own Sub/Function with parameters."),(0,o.kt)("p",null,"To create a Sub or Function with parameters, you simply add them to the definition line:"),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb"},"Sub WithSomeName(firstParameter as String)\n\nEnd Sub\n")),(0,o.kt)("p",null,"This approach is very simple. You give the parameter a name and a type declaration. This is very nice because it nearly exactly matches the ",(0,o.kt)("inlineCode",{parentName:"p"},"Dim")," statement with a Sub. That correspondence makes it very easy to start with an internally declared variable and then upgrade it to parameter. You can also go the other way: take a parameter and inline it into the Sub with some default or determined value. This is less common."),(0,o.kt)("p",null,"Once the parameter has been given a name and a type, you can simply use it within the Sub like any other variable. In this regard, your code will look the exact same. IF you are the person typing the VBA to use this Sub, then you will have to provide an appropriate variable as the parameter to make it all work."),(0,o.kt)("h3",{id:"declaring-an-optional-parameter"},"declaring an Optional parameter"),(0,o.kt)("p",null,"The one additional thing to consider is that of ",(0,o.kt)("inlineCode",{parentName:"p"},"Optional")," parameters. An optional parameter is one who is not strictly required. In lieu of a value, you can either leave the parameter missing or provide a default value. In either case, you can use the VBA specific function ",(0,o.kt)("inlineCode",{parentName:"p"},"IsMissing()")," to determine if the parameter was entered. An Optional parameter can be a very nice feature when you are trying to determine whether or not to make a Sub take parameters or just use defaults. You can provide the defaults in the parameter declaration and then allow the user (person typing the VBA) to override them if needed. This is a very common approach when writing library type code; provide sensible defaults that can be overwritten."))}d.isMDXComponent=!0}}]);