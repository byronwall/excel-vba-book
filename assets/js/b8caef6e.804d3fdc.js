"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[6658],{3905:function(e,t,a){a.d(t,{Zo:function(){return c},kt:function(){return m}});var n=a(7294);function r(e,t,a){return t in e?Object.defineProperty(e,t,{value:a,enumerable:!0,configurable:!0,writable:!0}):e[t]=a,e}function o(e,t){var a=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);t&&(n=n.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),a.push.apply(a,n)}return a}function i(e){for(var t=1;t<arguments.length;t++){var a=null!=arguments[t]?arguments[t]:{};t%2?o(Object(a),!0).forEach((function(t){r(e,t,a[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(a)):o(Object(a)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(a,t))}))}return e}function s(e,t){if(null==e)return{};var a,n,r=function(e,t){if(null==e)return{};var a,n,r={},o=Object.keys(e);for(n=0;n<o.length;n++)a=o[n],t.indexOf(a)>=0||(r[a]=e[a]);return r}(e,t);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);for(n=0;n<o.length;n++)a=o[n],t.indexOf(a)>=0||Object.prototype.propertyIsEnumerable.call(e,a)&&(r[a]=e[a])}return r}var l=n.createContext({}),u=function(e){var t=n.useContext(l),a=t;return e&&(a="function"==typeof e?e(t):i(i({},t),e)),a},c=function(e){var t=u(e.components);return n.createElement(l.Provider,{value:t},e.children)},h={inlineCode:"code",wrapper:function(e){var t=e.children;return n.createElement(n.Fragment,{},t)}},p=n.forwardRef((function(e,t){var a=e.components,r=e.mdxType,o=e.originalType,l=e.parentName,c=s(e,["components","mdxType","originalType","parentName"]),p=u(a),m=r,f=p["".concat(l,".").concat(m)]||p[m]||h[m]||o;return a?n.createElement(f,i(i({ref:t},c),{},{components:a})):n.createElement(f,i({ref:t},c))}));function m(e,t){var a=arguments,r=t&&t.mdxType;if("string"==typeof e||r){var o=a.length,i=new Array(o);i[0]=p;var s={};for(var l in t)hasOwnProperty.call(t,l)&&(s[l]=t[l]);s.originalType=e,s.mdxType="string"==typeof e?e:r,i[1]=s;for(var u=2;u<o;u++)i[u]=a[u];return n.createElement.apply(null,i)}return n.createElement.apply(null,a)}p.displayName="MDXCreateElement"},2468:function(e,t,a){a.r(t),a.d(t,{frontMatter:function(){return s},contentTitle:function(){return l},metadata:function(){return u},toc:function(){return c},default:function(){return p}});var n=a(7462),r=a(3366),o=(a(7294),a(3905)),i=["components"],s={},l=void 0,u={unversionedId:"overview-of-UDFs/13-05 managing-the-parameters-and-types-of-UDFs",id:"overview-of-UDFs/13-05 managing-the-parameters-and-types-of-UDFs",title:"13-05 managing-the-parameters-and-types-of-UDFs",description:"managing the parameters and types of UDFs",source:"@site/docs/13-overview-of-UDFs/13-05 managing-the-parameters-and-types-of-UDFs.md",sourceDirName:"13-overview-of-UDFs",slug:"/overview-of-UDFs/13-05 managing-the-parameters-and-types-of-UDFs",permalink:"/docs/overview-of-UDFs/13-05 managing-the-parameters-and-types-of-UDFs",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/13-overview-of-UDFs/13-05 managing-the-parameters-and-types-of-UDFs.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"13-04 limitations-of-UDfs",permalink:"/docs/overview-of-UDFs/13-04 limitations-of-UDfs"},next:{title:"13-06 complicated-UDFS",permalink:"/docs/overview-of-UDFs/13-06 complicated-UDFS"}},c=[{value:"managing the parameters and types of UDFs",id:"managing-the-parameters-and-types-of-udfs",children:[{value:"a note on return types",id:"a-note-on-return-types",children:[],level:3}],level:2}],h={toc:c};function p(e){var t=e.components,a=(0,r.Z)(e,i);return(0,o.kt)("wrapper",(0,n.Z)({},h,a,{components:t,mdxType:"MDXLayout"}),(0,o.kt)("h2",{id:"managing-the-parameters-and-types-of-udfs"},"managing the parameters and types of UDFs"),(0,o.kt)("p",null,"This section will focus on a topic that is quite nuanced but can have a large impact on how reusable your UDF code is. The focus here is on how to specify the type of the parameters and possibly the return of the UDF."),(0,o.kt)("p",null,"The reason things get tricky is that Excel is able to feed a wide range of object types to a UDF depending on how it was called. The common types to see are:"),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"Range"),(0,o.kt)("li",{parentName:"ul"},"Array/Variant"),(0,o.kt)("li",{parentName:"ul"},"Double/Number"),(0,o.kt)("li",{parentName:"ul"},"String"),(0,o.kt)("li",{parentName:"ul"},"Date"),(0,o.kt)("li",{parentName:"ul"},"Error")),(0,o.kt)("p",null,"The most common ways to call a UDF are"),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"Use a Range reference UDF(A1:B2)"),(0,o.kt)("li",{parentName:"ul"},"Use the result of some other operation UDF(5","*","A2). This can result in different object",(0,o.kt)("ul",{parentName:"li"},(0,o.kt)("li",{parentName:"ul"},"Array formula gives an array"),(0,o.kt)("li",{parentName:"ul"},"Math might give a number"),(0,o.kt)("li",{parentName:"ul"},"String formulas will give a string"),(0,o.kt)("li",{parentName:"ul"},"IF or CHOOSE might allow for multiple options depending on the result")))),(0,o.kt)("p",null,"Given this wide range of choices, it's important to consider how you intend for you UDF to be called and what types of inputs you want to be able to handle. You can choose to be as loose or as restrictive as you want on the parameter type, but this will have an impact on usage. If you go the loose route, you can call everything a Variant, but then you lose the utility of Intellisense as you are programming. If you go the strict route, you gain Intellisense, but might make your UDF fail on a simple case that it should be able to process."),(0,o.kt)("p",null,"As an example, let's say you've written a UDF that simple squares the number that it is fed. If you specify the parameter of this as a Range, your code will work fine with usages like UDF(A1), etc., but it will fail if someone sends in the result of math UDF(5","*","A1). This is odd because assuming that A1 is a number, there is no reason that you cannot square the result of that. Instead however, you will get an error that the result of that math (which is a Double) cannot be converted to a Range and your code will error out. For a simple example like this, it makes the most sense to declare the parameter as a Variant and just rely on the Value being correct."),(0,o.kt)("p",null,"TODO: add code for that example"),(0,o.kt)("p",null,"Things are fixed simple in that case, but it quickly becomes an issue when you want to handle different types of input. Maybe you are making a function that will concatenate an array of strings together. What happens when you only get a single string as a String instead of an Array containing Strings? Most likely, your code will fail in this instance, unless you've built int eh proper checks on the type. In this case, you will likely need to take a parameter of Variant and then do the checking to see how to handle it."),(0,o.kt)("p",null,"TODO: add an example of string concat code that works"),(0,o.kt)("p",null,"The most common spot to see this sort of issue is when deciding whether to deal with a type of Range or Variant (to handle an array). It is nice to work directly with Ranges and avoid the Variant, but this will make your code weak against someone who wants to use an array formula to call your UDF. It typically does no take much work to process an Array, but it helps to design things from th start like that."),(0,o.kt)("p",null,"TODO: add before example of UDF using Range"),(0,o.kt)("p",null,"TODO: add after example of that UDF using a Variant/Array instead of the Range"),(0,o.kt)("h3",{id:"a-note-on-return-types"},"a note on return types"),(0,o.kt)("p",null,"THe same thing can happen on the return side of the equation, but it is typically less of a problem. The main issues on the return side are returning arrays and dealing with Strings. If you want your UDF to work as an array formula, you can simply return an array and it will work. If that array is only a single cell, then it will look the same as a non-array formula."),(0,o.kt)("p",null,"Another issue is when working with Strings. If you return a string from a UDF, it will be formatted as Text instead of General. TODO: is that true? This can have intended consequences as Excel tends to treat Text differently when it is then sent to other functions. THe most common example is that a number stored as text will not be available for normal math operations."),(0,o.kt)("p",null,"You can avoid this by returning Variant but it can become an issue when you want a Function to work as a UDF and as a normal VBA Function. You might have a good reason to use a specific return type on the VBA side of things, but then Excel may not handle that the way you want (if using a String). Or, going the other way, you may have a UDF that works great because Excel can treat a single entry array as a single cell, but that becomes complicated when you call the UDF from another VBA location and then have to deal with a single number versus an array."))}p.isMDXComponent=!0}}]);