"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[6466],{3905:function(e,t,n){n.d(t,{Zo:function(){return c},kt:function(){return m}});var r=n(7294);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function a(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);t&&(r=r.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,r)}return n}function i(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?a(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):a(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function l(e,t){if(null==e)return{};var n,r,o=function(e,t){if(null==e)return{};var n,r,o={},a=Object.keys(e);for(r=0;r<a.length;r++)n=a[r],t.indexOf(n)>=0||(o[n]=e[n]);return o}(e,t);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(r=0;r<a.length;r++)n=a[r],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(o[n]=e[n])}return o}var u=r.createContext({}),s=function(e){var t=r.useContext(u),n=t;return e&&(n="function"==typeof e?e(t):i(i({},t),e)),n},c=function(e){var t=s(e.components);return r.createElement(u.Provider,{value:t},e.children)},p={inlineCode:"code",wrapper:function(e){var t=e.children;return r.createElement(r.Fragment,{},t)}},f=r.forwardRef((function(e,t){var n=e.components,o=e.mdxType,a=e.originalType,u=e.parentName,c=l(e,["components","mdxType","originalType","parentName"]),f=s(n),m=o,d=f["".concat(u,".").concat(m)]||f[m]||p[m]||a;return n?r.createElement(d,i(i({ref:t},c),{},{components:n})):r.createElement(d,i({ref:t},c))}));function m(e,t){var n=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var a=n.length,i=new Array(a);i[0]=f;var l={};for(var u in t)hasOwnProperty.call(t,u)&&(l[u]=t[u]);l.originalType=e,l.mdxType="string"==typeof e?e:o,i[1]=l;for(var s=2;s<a;s++)i[s]=n[s];return r.createElement.apply(null,i)}return r.createElement.apply(null,n)}f.displayName="MDXCreateElement"},5129:function(e,t,n){n.r(t),n.d(t,{frontMatter:function(){return l},contentTitle:function(){return u},metadata:function(){return s},toc:function(){return c},default:function(){return f}});var r=n(7462),o=n(3366),a=(n(7294),n(3905)),i=["components"],l={},u=void 0,s={unversionedId:"overview-of-values-and-formulas/05-13 SplitIntoRowsmd",id:"overview-of-values-and-formulas/05-13 SplitIntoRowsmd",title:"05-13 SplitIntoRowsmd",description:"SplitIntoRows.md",source:"@site/docs/05-overview-of-values-and-formulas/05-13 SplitIntoRowsmd.md",sourceDirName:"05-overview-of-values-and-formulas",slug:"/overview-of-values-and-formulas/05-13 SplitIntoRowsmd",permalink:"/excel-vba-book/docs/overview-of-values-and-formulas/05-13 SplitIntoRowsmd",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/05-overview-of-values-and-formulas/05-13 SplitIntoRowsmd.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"05-13 MakeHyperlinksmd",permalink:"/excel-vba-book/docs/overview-of-values-and-formulas/05-13 MakeHyperlinksmd"},next:{title:"05-14 TrimSelectionmd",permalink:"/excel-vba-book/docs/overview-of-values-and-formulas/05-14 TrimSelectionmd"}},c=[{value:"SplitIntoRows.md",id:"splitintorowsmd",children:[],level:2}],p={toc:c};function f(e){var t=e.components,n=(0,o.Z)(e,i);return(0,a.kt)("wrapper",(0,r.Z)({},p,n,{components:t,mdxType:"MDXLayout"}),(0,a.kt)("h2",{id:"splitintorowsmd"},"SplitIntoRows.md"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb"},'Public Sub SplitIntoRows()\n\n    Dim outputRange As Range\n\n    Dim inputRange As Range\n    Set inputRange = Selection\n\n    Set outputRange = GetInputOrSelection("Select the output corner")\n\n    Dim targetPart As Variant\n    Dim offsetCounter As Long\n    offsetCounter = 0\n    Dim targetCell As Range\n\n    For Each targetCell In inputRange.SpecialCells(xlCellTypeVisible)\n        Dim targetParts As Variant\n        targetParts = Split(targetCell, vbLf)\n\n        For Each targetPart In targetParts\n            outputRange.Offset(offsetCounter) = targetPart\n\n            offsetCounter = offsetCounter + 1\n        Next targetPart\n    Next targetCell\nEnd Sub\n')))}f.isMDXComponent=!0}}]);