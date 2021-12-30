"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[1986],{3905:function(e,n,t){t.d(n,{Zo:function(){return a},kt:function(){return p}});var r=t(7294);function o(e,n,t){return n in e?Object.defineProperty(e,n,{value:t,enumerable:!0,configurable:!0,writable:!0}):e[n]=t,e}function i(e,n){var t=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);n&&(r=r.filter((function(n){return Object.getOwnPropertyDescriptor(e,n).enumerable}))),t.push.apply(t,r)}return t}function l(e){for(var n=1;n<arguments.length;n++){var t=null!=arguments[n]?arguments[n]:{};n%2?i(Object(t),!0).forEach((function(n){o(e,n,t[n])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(t)):i(Object(t)).forEach((function(n){Object.defineProperty(e,n,Object.getOwnPropertyDescriptor(t,n))}))}return e}function u(e,n){if(null==e)return{};var t,r,o=function(e,n){if(null==e)return{};var t,r,o={},i=Object.keys(e);for(r=0;r<i.length;r++)t=i[r],n.indexOf(t)>=0||(o[t]=e[t]);return o}(e,n);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(r=0;r<i.length;r++)t=i[r],n.indexOf(t)>=0||Object.prototype.propertyIsEnumerable.call(e,t)&&(o[t]=e[t])}return o}var c=r.createContext({}),d=function(e){var n=r.useContext(c),t=n;return e&&(t="function"==typeof e?e(n):l(l({},n),e)),t},a=function(e){var n=d(e.components);return r.createElement(c.Provider,{value:n},e.children)},s={inlineCode:"code",wrapper:function(e){var n=e.children;return r.createElement(r.Fragment,{},n)}},m=r.forwardRef((function(e,n){var t=e.components,o=e.mdxType,i=e.originalType,c=e.parentName,a=u(e,["components","mdxType","originalType","parentName"]),m=d(t),p=o,f=m["".concat(c,".").concat(p)]||m[p]||s[p]||i;return t?r.createElement(f,l(l({ref:n},a),{},{components:t})):r.createElement(f,l({ref:n},a))}));function p(e,n){var t=arguments,o=n&&n.mdxType;if("string"==typeof e||o){var i=t.length,l=new Array(i);l[0]=m;var u={};for(var c in n)hasOwnProperty.call(n,c)&&(u[c]=n[c]);u.originalType=e,u.mdxType="string"==typeof e?e:o,l[1]=u;for(var d=2;d<i;d++)l[d]=t[d];return r.createElement.apply(null,l)}return r.createElement.apply(null,t)}m.displayName="MDXCreateElement"},4581:function(e,n,t){t.r(n),t.d(n,{frontMatter:function(){return u},contentTitle:function(){return c},metadata:function(){return d},toc:function(){return a},default:function(){return m}});var r=t(7462),o=t(3366),i=(t(7294),t(3905)),l=["components"],u={},c=void 0,d={unversionedId:"overview-of-utility-code/15-14 UnhideAllRowsAndColumnsmd",id:"overview-of-utility-code/15-14 UnhideAllRowsAndColumnsmd",title:"15-14 UnhideAllRowsAndColumnsmd",description:"UnhideAllRowsAndColumns.md",source:"@site/docs/15-overview-of-utility-code/15-14 UnhideAllRowsAndColumnsmd.md",sourceDirName:"15-overview-of-utility-code",slug:"/overview-of-utility-code/15-14 UnhideAllRowsAndColumnsmd",permalink:"/docs/overview-of-utility-code/15-14 UnhideAllRowsAndColumnsmd",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/15-overview-of-utility-code/15-14 UnhideAllRowsAndColumnsmd.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"15-13 SheetDeleteHiddenRowsmd",permalink:"/docs/overview-of-utility-code/15-13 SheetDeleteHiddenRowsmd"},next:{title:"intro",permalink:"/docs/intro"}},a=[{value:"UnhideAllRowsAndColumns.md",id:"unhideallrowsandcolumnsmd",children:[],level:2}],s={toc:a};function m(e){var n=e.components,t=(0,o.Z)(e,l);return(0,i.kt)("wrapper",(0,r.Z)({},s,t,{components:n,mdxType:"MDXLayout"}),(0,i.kt)("h2",{id:"unhideallrowsandcolumnsmd"},"UnhideAllRowsAndColumns.md"),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-vb"},"Public Sub UnhideAllRowsAndColumns()\n\n    ActiveSheet.Cells.EntireRow.Hidden = False\n    ActiveSheet.Cells.EntireColumn.Hidden = False\n\nEnd Sub\n")))}m.isMDXComponent=!0}}]);