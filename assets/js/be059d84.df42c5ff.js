"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[6720],{3905:function(e,t,n){n.d(t,{Zo:function(){return c},kt:function(){return d}});var r=n(7294);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function a(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);t&&(r=r.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,r)}return n}function l(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?a(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):a(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function i(e,t){if(null==e)return{};var n,r,o=function(e,t){if(null==e)return{};var n,r,o={},a=Object.keys(e);for(r=0;r<a.length;r++)n=a[r],t.indexOf(n)>=0||(o[n]=e[n]);return o}(e,t);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(r=0;r<a.length;r++)n=a[r],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(o[n]=e[n])}return o}var u=r.createContext({}),s=function(e){var t=r.useContext(u),n=t;return e&&(n="function"==typeof e?e(t):l(l({},t),e)),n},c=function(e){var t=s(e.components);return r.createElement(u.Provider,{value:t},e.children)},p={inlineCode:"code",wrapper:function(e){var t=e.children;return r.createElement(r.Fragment,{},t)}},m=r.forwardRef((function(e,t){var n=e.components,o=e.mdxType,a=e.originalType,u=e.parentName,c=i(e,["components","mdxType","originalType","parentName"]),m=s(n),d=o,f=m["".concat(u,".").concat(d)]||m[d]||p[d]||a;return n?r.createElement(f,l(l({ref:t},c),{},{components:n})):r.createElement(f,l({ref:t},c))}));function d(e,t){var n=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var a=n.length,l=new Array(a);l[0]=m;var i={};for(var u in t)hasOwnProperty.call(t,u)&&(i[u]=t[u]);i.originalType=e,i.mdxType="string"==typeof e?e:o,l[1]=i;for(var s=2;s<a;s++)l[s]=n[s];return r.createElement.apply(null,l)}return r.createElement.apply(null,n)}m.displayName="MDXCreateElement"},2969:function(e,t,n){n.r(t),n.d(t,{frontMatter:function(){return i},contentTitle:function(){return u},metadata:function(){return s},toc:function(){return c},default:function(){return m}});var r=n(7462),o=n(3366),a=(n(7294),n(3905)),l=["components"],i={},u=void 0,s={unversionedId:"overview-of-values-and-formulas/05-18 SplitIntoColumnsmd",id:"overview-of-values-and-formulas/05-18 SplitIntoColumnsmd",title:"05-18 SplitIntoColumnsmd",description:"SplitIntoColumns.md",source:"@site/docs/05-overview-of-values-and-formulas/05-18 SplitIntoColumnsmd.md",sourceDirName:"05-overview-of-values-and-formulas",slug:"/overview-of-values-and-formulas/05-18 SplitIntoColumnsmd",permalink:"/docs/overview-of-values-and-formulas/05-18 SplitIntoColumnsmd",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/05-overview-of-values-and-formulas/05-18 SplitIntoColumnsmd.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"05-17 SplitAndKeepmd",permalink:"/docs/overview-of-values-and-formulas/05-17 SplitAndKeepmd"},next:{title:"05-19 SplitIntoRowsmd",permalink:"/docs/overview-of-values-and-formulas/05-19 SplitIntoRowsmd"}},c=[{value:"SplitIntoColumns.md",id:"splitintocolumnsmd",children:[],level:2}],p={toc:c};function m(e){var t=e.components,n=(0,o.Z)(e,l);return(0,a.kt)("wrapper",(0,r.Z)({},p,n,{components:t,mdxType:"MDXLayout"}),(0,a.kt)("h2",{id:"splitintocolumnsmd"},"SplitIntoColumns.md"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb"},'Public Sub SplitIntoColumns()\n\n    Dim inputRange As Range\n\n    Set inputRange = GetInputOrSelection("Select the range of cells to split")\n\n    Dim targetCell As Range\n\n    Dim delimiter As String\n    delimiter = Application.InputBox("What is the delimiter?", , ",", vbOKCancel)\n    If delimiter = "" Or delimiter = "False" Then GoTo errHandler\n    For Each targetCell In inputRange\n\n        Dim targetCellParts As Variant\n        targetCellParts = Split(targetCell, delimiter)\n\n        Dim targetPart As Variant\n        For Each targetPart In targetCellParts\n\n            Set targetCell = targetCell.Offset(, 1)\n            targetCell = targetPart\n\n        Next targetPart\n\n    Next targetCell\n    Exit Sub\nerrHandler:\n    MsgBox "No Delimiter Defined!"\nEnd Sub\n')))}m.isMDXComponent=!0}}]);