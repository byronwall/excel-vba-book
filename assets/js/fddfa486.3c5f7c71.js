"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[8486],{3905:function(e,t,r){r.d(t,{Zo:function(){return u},kt:function(){return d}});var o=r(7294);function n(e,t,r){return t in e?Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}):e[t]=r,e}function i(e,t){var r=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);t&&(o=o.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),r.push.apply(r,o)}return r}function a(e){for(var t=1;t<arguments.length;t++){var r=null!=arguments[t]?arguments[t]:{};t%2?i(Object(r),!0).forEach((function(t){n(e,t,r[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(r)):i(Object(r)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(r,t))}))}return e}function s(e,t){if(null==e)return{};var r,o,n=function(e,t){if(null==e)return{};var r,o,n={},i=Object.keys(e);for(o=0;o<i.length;o++)r=i[o],t.indexOf(r)>=0||(n[r]=e[r]);return n}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(o=0;o<i.length;o++)r=i[o],t.indexOf(r)>=0||Object.prototype.propertyIsEnumerable.call(e,r)&&(n[r]=e[r])}return n}var c=o.createContext({}),h=function(e){var t=o.useContext(c),r=t;return e&&(r="function"==typeof e?e(t):a(a({},t),e)),r},u=function(e){var t=h(e.components);return o.createElement(c.Provider,{value:t},e.children)},l={inlineCode:"code",wrapper:function(e){var t=e.children;return o.createElement(o.Fragment,{},t)}},p=o.forwardRef((function(e,t){var r=e.components,n=e.mdxType,i=e.originalType,c=e.parentName,u=s(e,["components","mdxType","originalType","parentName"]),p=h(r),d=n,f=p["".concat(c,".").concat(d)]||p[d]||l[d]||i;return r?o.createElement(f,a(a({ref:t},u),{},{components:r})):o.createElement(f,a({ref:t},u))}));function d(e,t){var r=arguments,n=t&&t.mdxType;if("string"==typeof e||n){var i=r.length,a=new Array(i);a[0]=p;var s={};for(var c in t)hasOwnProperty.call(t,c)&&(s[c]=t[c]);s.originalType=e,s.mdxType="string"==typeof e?e:n,a[1]=s;for(var h=2;h<i;h++)a[h]=r[h];return o.createElement.apply(null,a)}return o.createElement.apply(null,r)}p.displayName="MDXCreateElement"},6241:function(e,t,r){r.r(t),r.d(t,{frontMatter:function(){return s},contentTitle:function(){return c},metadata:function(){return h},toc:function(){return u},default:function(){return p}});var o=r(7462),n=r(3366),i=(r(7294),r(3905)),a=["components"],s={},c=void 0,h={unversionedId:"The-Worksheet-object/07-01 introduction-to-the-Worksheet-object",id:"The-Worksheet-object/07-01 introduction-to-the-Worksheet-object",title:"07-01 introduction-to-the-Worksheet-object",description:"introduction to the Worksheet object",source:"@site/docs/07-The-Worksheet-object/07-01 introduction-to-the-Worksheet-object.md",sourceDirName:"07-The-Worksheet-object",slug:"/The-Worksheet-object/07-01 introduction-to-the-Worksheet-object",permalink:"/docs/The-Worksheet-object/07-01 introduction-to-the-Worksheet-object",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/07-The-Worksheet-object/07-01 introduction-to-the-Worksheet-object.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"The Worksheet object",permalink:"/docs/The-Worksheet-object/07 The-Worksheet-object"},next:{title:"07-02 creating-and-managing-Worksheets",permalink:"/docs/The-Worksheet-object/07-02 creating-and-managing-Worksheets"}},u=[{value:"introduction to the Worksheet object",id:"introduction-to-the-worksheet-object",children:[],level:2}],l={toc:u};function p(e){var t=e.components,r=(0,n.Z)(e,a);return(0,i.kt)("wrapper",(0,o.Z)({},l,r,{components:t,mdxType:"MDXLayout"}),(0,i.kt)("h2",{id:"introduction-to-the-worksheet-object"},"introduction to the Worksheet object"),(0,i.kt)("p",null,"This chapter will focus on the aspects of the Worksheet that appear commonly in VBA code. This chapter is a little shorter than others because in general, the Worksheet is a conduit to more useful things. There is very little that takes place within the Worksheet object that is not just a pass through to the more interesting details (e.g. Range or Chart). Having said that, there are a handful of areas that are relevant to the Worksheet and not accessible anywhere else. Those specific areas include:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"Creating and managing Worksheets -- this sounds obvious but managing the references to Worksheets becomes a major issue when working with large, complicated workflows"),(0,i.kt)("li",{parentName:"ul"},"Print layout, printing, and exporting"),(0,i.kt)("li",{parentName:"ul"},"Locking and setting passwords on Worksheets"),(0,i.kt)("li",{parentName:"ul"},"Managing the properties of the Worksheet itself including Name, tab color, etc.")),(0,i.kt)("p",null,"TODO: any other Worksheet things?"),(0,i.kt)("p",null,"Of the topics listed above, the most important area is actually creating and managing the Worksheets in a complicated workflow. This is closely related to working with Ranges since presumably you create the Worksheet to put data into or something else into it. Managing the references to Worksheets can be a big deal and determining how best to access or select a given Worksheet can be important. In addition to getting references, there are a handful of times where you actually need to Activate a Worksheet. Knowing when this is and is not required is important."),(0,i.kt)("p",null,"TODO: when do you have to Activate?"))}p.isMDXComponent=!0}}]);