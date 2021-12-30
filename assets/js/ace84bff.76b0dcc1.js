"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[973],{3905:function(e,n,r){r.d(n,{Zo:function(){return s},kt:function(){return m}});var t=r(7294);function o(e,n,r){return n in e?Object.defineProperty(e,n,{value:r,enumerable:!0,configurable:!0,writable:!0}):e[n]=r,e}function a(e,n){var r=Object.keys(e);if(Object.getOwnPropertySymbols){var t=Object.getOwnPropertySymbols(e);n&&(t=t.filter((function(n){return Object.getOwnPropertyDescriptor(e,n).enumerable}))),r.push.apply(r,t)}return r}function c(e){for(var n=1;n<arguments.length;n++){var r=null!=arguments[n]?arguments[n]:{};n%2?a(Object(r),!0).forEach((function(n){o(e,n,r[n])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(r)):a(Object(r)).forEach((function(n){Object.defineProperty(e,n,Object.getOwnPropertyDescriptor(r,n))}))}return e}function i(e,n){if(null==e)return{};var r,t,o=function(e,n){if(null==e)return{};var r,t,o={},a=Object.keys(e);for(t=0;t<a.length;t++)r=a[t],n.indexOf(r)>=0||(o[r]=e[r]);return o}(e,n);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(t=0;t<a.length;t++)r=a[t],n.indexOf(r)>=0||Object.prototype.propertyIsEnumerable.call(e,r)&&(o[r]=e[r])}return o}var l=t.createContext({}),u=function(e){var n=t.useContext(l),r=n;return e&&(r="function"==typeof e?e(n):c(c({},n),e)),r},s=function(e){var n=u(e.components);return t.createElement(l.Provider,{value:n},e.children)},d={inlineCode:"code",wrapper:function(e){var n=e.children;return t.createElement(t.Fragment,{},n)}},p=t.forwardRef((function(e,n){var r=e.components,o=e.mdxType,a=e.originalType,l=e.parentName,s=i(e,["components","mdxType","originalType","parentName"]),p=u(r),m=o,f=p["".concat(l,".").concat(m)]||p[m]||d[m]||a;return r?t.createElement(f,c(c({ref:n},s),{},{components:r})):t.createElement(f,c({ref:n},s))}));function m(e,n){var r=arguments,o=n&&n.mdxType;if("string"==typeof e||o){var a=r.length,c=new Array(a);c[0]=p;var i={};for(var l in n)hasOwnProperty.call(n,l)&&(i[l]=n[l]);i.originalType=e,i.mdxType="string"==typeof e?e:o,c[1]=i;for(var u=2;u<a;u++)c[u]=r[u];return t.createElement.apply(null,c)}return t.createElement.apply(null,r)}p.displayName="MDXCreateElement"},7146:function(e,n,r){r.r(n),r.d(n,{frontMatter:function(){return i},contentTitle:function(){return l},metadata:function(){return u},toc:function(){return s},default:function(){return p}});var t=r(7462),o=r(3366),a=(r(7294),r(3905)),c=["components"],i={},l=void 0,u={unversionedId:"overview-of-UDFs/13-09 ConcatRangemd",id:"overview-of-UDFs/13-09 ConcatRangemd",title:"13-09 ConcatRangemd",description:"ConcatRange.md",source:"@site/docs/13-overview-of-UDFs/13-09 ConcatRangemd.md",sourceDirName:"13-overview-of-UDFs",slug:"/overview-of-UDFs/13-09 ConcatRangemd",permalink:"/docs/overview-of-UDFs/13-09 ConcatRangemd",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/13-overview-of-UDFs/13-09 ConcatRangemd.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"13-08 ConcatArrmd",permalink:"/docs/overview-of-UDFs/13-08 ConcatArrmd"},next:{title:"13-10 RandLettersmd",permalink:"/docs/overview-of-UDFs/13-10 RandLettersmd"}},s=[{value:"ConcatRange.md",id:"concatrangemd",children:[],level:2}],d={toc:s};function p(e){var n=e.components,r=(0,o.Z)(e,c);return(0,a.kt)("wrapper",(0,t.Z)({},d,r,{components:n,mdxType:"MDXLayout"}),(0,a.kt)("h2",{id:"concatrangemd"},"ConcatRange.md"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb"},"Public Function ConcatRange(rngCells As Range, strDelim As String) As String\n    Dim cellCount As Long\n\n    cellCount = rngCells.CountLarge\n\n    Dim arrValues As Variant\n    ReDim arrValues(1 To cellCount)\n\n    Dim index As Long\n    index = 1\n\n    Dim rngCell As Range\n    For Each rngCell In rngCells\n        arrValues(index) = rngCell\n\n        index = index + 1\n    Next\n\n    ConcatRange = Join(arrValues, strDelim)\nEnd Function\n")))}p.isMDXComponent=!0}}]);