"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[9875],{3905:function(e,r,n){n.d(r,{Zo:function(){return s},kt:function(){return f}});var t=n(7294);function o(e,r,n){return r in e?Object.defineProperty(e,r,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[r]=n,e}function a(e,r){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var t=Object.getOwnPropertySymbols(e);r&&(t=t.filter((function(r){return Object.getOwnPropertyDescriptor(e,r).enumerable}))),n.push.apply(n,t)}return n}function c(e){for(var r=1;r<arguments.length;r++){var n=null!=arguments[r]?arguments[r]:{};r%2?a(Object(n),!0).forEach((function(r){o(e,r,n[r])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):a(Object(n)).forEach((function(r){Object.defineProperty(e,r,Object.getOwnPropertyDescriptor(n,r))}))}return e}function i(e,r){if(null==e)return{};var n,t,o=function(e,r){if(null==e)return{};var n,t,o={},a=Object.keys(e);for(t=0;t<a.length;t++)n=a[t],r.indexOf(n)>=0||(o[n]=e[n]);return o}(e,r);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(t=0;t<a.length;t++)n=a[t],r.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(o[n]=e[n])}return o}var u=t.createContext({}),l=function(e){var r=t.useContext(u),n=r;return e&&(n="function"==typeof e?e(r):c(c({},r),e)),n},s=function(e){var r=l(e.components);return t.createElement(u.Provider,{value:r},e.children)},p={inlineCode:"code",wrapper:function(e){var r=e.children;return t.createElement(t.Fragment,{},r)}},d=t.forwardRef((function(e,r){var n=e.components,o=e.mdxType,a=e.originalType,u=e.parentName,s=i(e,["components","mdxType","originalType","parentName"]),d=l(n),f=o,m=d["".concat(u,".").concat(f)]||d[f]||p[f]||a;return n?t.createElement(m,c(c({ref:r},s),{},{components:n})):t.createElement(m,c({ref:r},s))}));function f(e,r){var n=arguments,o=r&&r.mdxType;if("string"==typeof e||o){var a=n.length,c=new Array(a);c[0]=d;var i={};for(var u in r)hasOwnProperty.call(r,u)&&(i[u]=r[u]);i.originalType=e,i.mdxType="string"==typeof e?e:o,c[1]=i;for(var l=2;l<a;l++)c[l]=n[l];return t.createElement.apply(null,c)}return t.createElement.apply(null,n)}d.displayName="MDXCreateElement"},5363:function(e,r,n){n.r(r),n.d(r,{frontMatter:function(){return i},contentTitle:function(){return u},metadata:function(){return l},toc:function(){return s},default:function(){return d}});var t=n(7462),o=n(3366),a=(n(7294),n(3905)),c=["components"],i={},u=void 0,l={unversionedId:"overview-of-UDFs/13-08 ConcatArrmd",id:"overview-of-UDFs/13-08 ConcatArrmd",title:"13-08 ConcatArrmd",description:"ConcatArr.md",source:"@site/docs/13-overview-of-UDFs/13-08 ConcatArrmd.md",sourceDirName:"13-overview-of-UDFs",slug:"/overview-of-UDFs/13-08 ConcatArrmd",permalink:"/docs/overview-of-UDFs/13-08 ConcatArrmd",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/13-overview-of-UDFs/13-08 ConcatArrmd.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"13-07 debugging-UDFs",permalink:"/docs/overview-of-UDFs/13-07 debugging-UDFs"},next:{title:"13-09 ConcatRangemd",permalink:"/docs/overview-of-UDFs/13-09 ConcatRangemd"}},s=[{value:"ConcatArr.md",id:"concatarrmd",children:[],level:2}],p={toc:s};function d(e){var r=e.components,n=(0,o.Z)(e,c);return(0,a.kt)("wrapper",(0,t.Z)({},p,n,{components:r,mdxType:"MDXLayout"}),(0,a.kt)("h2",{id:"concatarrmd"},"ConcatArr.md"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb"},"Public Function ConcatArr(rngCells As Variant, strDelim As String) As String\n    Dim cellCount As Long\n\n    cellCount = UBound(rngCells, 1)\n\n    Dim arrValues As Variant\n    ReDim arrValues(1 To cellCount)\n\n    Dim index As Long\n    index = 1\n\n    Dim rngCell As Variant\n    For Each rngCell In rngCells\n        arrValues(index) = rngCell\n\n        index = index + 1\n    Next\n\n    ConcatArr = Join(arrValues, strDelim)\nEnd Function\n")))}d.isMDXComponent=!0}}]);