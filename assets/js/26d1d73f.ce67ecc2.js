"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[2022],{3905:function(e,t,r){r.d(t,{Zo:function(){return d},kt:function(){return m}});var n=r(7294);function o(e,t,r){return t in e?Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}):e[t]=r,e}function c(e,t){var r=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);t&&(n=n.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),r.push.apply(r,n)}return r}function i(e){for(var t=1;t<arguments.length;t++){var r=null!=arguments[t]?arguments[t]:{};t%2?c(Object(r),!0).forEach((function(t){o(e,t,r[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(r)):c(Object(r)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(r,t))}))}return e}function a(e,t){if(null==e)return{};var r,n,o=function(e,t){if(null==e)return{};var r,n,o={},c=Object.keys(e);for(n=0;n<c.length;n++)r=c[n],t.indexOf(r)>=0||(o[r]=e[r]);return o}(e,t);if(Object.getOwnPropertySymbols){var c=Object.getOwnPropertySymbols(e);for(n=0;n<c.length;n++)r=c[n],t.indexOf(r)>=0||Object.prototype.propertyIsEnumerable.call(e,r)&&(o[r]=e[r])}return o}var l=n.createContext({}),u=function(e){var t=n.useContext(l),r=t;return e&&(r="function"==typeof e?e(t):i(i({},t),e)),r},d=function(e){var t=u(e.components);return n.createElement(l.Provider,{value:t},e.children)},p={inlineCode:"code",wrapper:function(e){var t=e.children;return n.createElement(n.Fragment,{},t)}},f=n.forwardRef((function(e,t){var r=e.components,o=e.mdxType,c=e.originalType,l=e.parentName,d=a(e,["components","mdxType","originalType","parentName"]),f=u(r),m=o,s=f["".concat(l,".").concat(m)]||f[m]||p[m]||c;return r?n.createElement(s,i(i({ref:t},d),{},{components:r})):n.createElement(s,i({ref:t},d))}));function m(e,t){var r=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var c=r.length,i=new Array(c);i[0]=f;var a={};for(var l in t)hasOwnProperty.call(t,l)&&(a[l]=t[l]);a.originalType=e,a.mdxType="string"==typeof e?e:o,i[1]=a;for(var u=2;u<c;u++)i[u]=r[u];return n.createElement.apply(null,i)}return n.createElement.apply(null,r)}f.displayName="MDXCreateElement"},2394:function(e,t,r){r.r(t),r.d(t,{frontMatter:function(){return a},contentTitle:function(){return l},metadata:function(){return u},toc:function(){return d},default:function(){return f}});var n=r(7462),o=r(3366),c=(r(7294),r(3905)),i=["components"],a={},l=void 0,u={unversionedId:"overview-of-utility-code/15-07 ForceRecalcmd",id:"overview-of-utility-code/15-07 ForceRecalcmd",title:"15-07 ForceRecalcmd",description:"ForceRecalc.md",source:"@site/docs/15-overview-of-utility-code/15-07 ForceRecalcmd.md",sourceDirName:"15-overview-of-utility-code",slug:"/overview-of-utility-code/15-07 ForceRecalcmd",permalink:"/docs/overview-of-utility-code/15-07 ForceRecalcmd",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/15-overview-of-utility-code/15-07 ForceRecalcmd.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"15-06 FillValueDownmd",permalink:"/docs/overview-of-utility-code/15-06 FillValueDownmd"},next:{title:"15-08 GenerateRandomDatamd",permalink:"/docs/overview-of-utility-code/15-08 GenerateRandomDatamd"}},d=[{value:"ForceRecalc.md",id:"forcerecalcmd",children:[],level:2}],p={toc:d};function f(e){var t=e.components,r=(0,o.Z)(e,i);return(0,c.kt)("wrapper",(0,n.Z)({},p,r,{components:t,mdxType:"MDXLayout"}),(0,c.kt)("h2",{id:"forcerecalcmd"},"ForceRecalc.md"),(0,c.kt)("pre",null,(0,c.kt)("code",{parentName:"pre",className:"language-vb"},"Public Sub ForceRecalc()\n\n    Application.CalculateFullRebuild\n\nEnd Sub\n")))}f.isMDXComponent=!0}}]);