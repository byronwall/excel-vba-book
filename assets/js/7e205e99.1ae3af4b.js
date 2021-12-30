"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[7240],{3905:function(e,t,r){r.d(t,{Zo:function(){return d},kt:function(){return v}});var n=r(7294);function o(e,t,r){return t in e?Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}):e[t]=r,e}function i(e,t){var r=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);t&&(n=n.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),r.push.apply(r,n)}return r}function l(e){for(var t=1;t<arguments.length;t++){var r=null!=arguments[t]?arguments[t]:{};t%2?i(Object(r),!0).forEach((function(t){o(e,t,r[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(r)):i(Object(r)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(r,t))}))}return e}function a(e,t){if(null==e)return{};var r,n,o=function(e,t){if(null==e)return{};var r,n,o={},i=Object.keys(e);for(n=0;n<i.length;n++)r=i[n],t.indexOf(r)>=0||(o[r]=e[r]);return o}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(n=0;n<i.length;n++)r=i[n],t.indexOf(r)>=0||Object.prototype.propertyIsEnumerable.call(e,r)&&(o[r]=e[r])}return o}var c=n.createContext({}),u=function(e){var t=n.useContext(c),r=t;return e&&(r="function"==typeof e?e(t):l(l({},t),e)),r},d=function(e){var t=u(e.components);return n.createElement(c.Provider,{value:t},e.children)},s={inlineCode:"code",wrapper:function(e){var t=e.children;return n.createElement(n.Fragment,{},t)}},p=n.forwardRef((function(e,t){var r=e.components,o=e.mdxType,i=e.originalType,c=e.parentName,d=a(e,["components","mdxType","originalType","parentName"]),p=u(r),v=o,f=p["".concat(c,".").concat(v)]||p[v]||s[v]||i;return r?n.createElement(f,l(l({ref:t},d),{},{components:r})):n.createElement(f,l({ref:t},d))}));function v(e,t){var r=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var i=r.length,l=new Array(i);l[0]=p;var a={};for(var c in t)hasOwnProperty.call(t,c)&&(a[c]=t[c]);a.originalType=e,a.mdxType="string"==typeof e?e:o,l[1]=a;for(var u=2;u<i;u++)l[u]=r[u];return n.createElement.apply(null,l)}return n.createElement.apply(null,r)}p.displayName="MDXCreateElement"},1786:function(e,t,r){r.r(t),r.d(t,{frontMatter:function(){return a},contentTitle:function(){return c},metadata:function(){return u},toc:function(){return d},default:function(){return p}});var n=r(7462),o=r(3366),i=(r(7294),r(3905)),l=["components"],a={},c=void 0,u={unversionedId:"overview-of-utility-code/15-10 PivotSetAllFieldsmd",id:"overview-of-utility-code/15-10 PivotSetAllFieldsmd",title:"15-10 PivotSetAllFieldsmd",description:"PivotSetAllFields.md",source:"@site/docs/15-overview-of-utility-code/15-10 PivotSetAllFieldsmd.md",sourceDirName:"15-overview-of-utility-code",slug:"/overview-of-utility-code/15-10 PivotSetAllFieldsmd",permalink:"/excel-vba-book/docs/overview-of-utility-code/15-10 PivotSetAllFieldsmd",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/15-overview-of-utility-code/15-10 PivotSetAllFieldsmd.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"15-09 OpenContainingFoldermd",permalink:"/excel-vba-book/docs/overview-of-utility-code/15-09 OpenContainingFoldermd"},next:{title:"15-11 SeriesSplitmd",permalink:"/excel-vba-book/docs/overview-of-utility-code/15-11 SeriesSplitmd"}},d=[{value:"PivotSetAllFields.md",id:"pivotsetallfieldsmd",children:[],level:2}],s={toc:d};function p(e){var t=e.components,r=(0,o.Z)(e,l);return(0,i.kt)("wrapper",(0,n.Z)({},s,r,{components:t,mdxType:"MDXLayout"}),(0,i.kt)("h2",{id:"pivotsetallfieldsmd"},"PivotSetAllFields.md"),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-vb"},'Public Sub PivotSetAllFields()\n\n    Dim targetTable As PivotTable\n    Dim targetSheet As Worksheet\n\n    Set targetSheet = ActiveSheet\n\n    \'this information is a bit unclear to me\n    MsgBox "This defaults to the average for every Pivot table on the sheet.  Edit code for other result."\n    On Error Resume Next\n    For Each targetTable In targetSheet.PivotTables\n        Dim targetField As PivotField\n        For Each targetField In targetTable.DataFields\n            targetField.Function = xlAverage\n        Next targetField\n    Next targetTable\n\nEnd Sub\n')))}p.isMDXComponent=!0}}]);