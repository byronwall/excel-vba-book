"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[3592],{3905:function(e,t,r){r.d(t,{Zo:function(){return u},kt:function(){return m}});var n=r(7294);function o(e,t,r){return t in e?Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}):e[t]=r,e}function a(e,t){var r=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);t&&(n=n.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),r.push.apply(r,n)}return r}function i(e){for(var t=1;t<arguments.length;t++){var r=null!=arguments[t]?arguments[t]:{};t%2?a(Object(r),!0).forEach((function(t){o(e,t,r[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(r)):a(Object(r)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(r,t))}))}return e}function c(e,t){if(null==e)return{};var r,n,o=function(e,t){if(null==e)return{};var r,n,o={},a=Object.keys(e);for(n=0;n<a.length;n++)r=a[n],t.indexOf(r)>=0||(o[r]=e[r]);return o}(e,t);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(n=0;n<a.length;n++)r=a[n],t.indexOf(r)>=0||Object.prototype.propertyIsEnumerable.call(e,r)&&(o[r]=e[r])}return o}var l=n.createContext({}),p=function(e){var t=n.useContext(l),r=t;return e&&(r="function"==typeof e?e(t):i(i({},t),e)),r},u=function(e){var t=p(e.components);return n.createElement(l.Provider,{value:t},e.children)},d={inlineCode:"code",wrapper:function(e){var t=e.children;return n.createElement(n.Fragment,{},t)}},s=n.forwardRef((function(e,t){var r=e.components,o=e.mdxType,a=e.originalType,l=e.parentName,u=c(e,["components","mdxType","originalType","parentName"]),s=p(r),m=o,f=s["".concat(l,".").concat(m)]||s[m]||d[m]||a;return r?n.createElement(f,i(i({ref:t},u),{},{components:r})):n.createElement(f,i({ref:t},u))}));function m(e,t){var r=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var a=r.length,i=new Array(a);i[0]=s;var c={};for(var l in t)hasOwnProperty.call(t,l)&&(c[l]=t[l]);c.originalType=e,c.mdxType="string"==typeof e?e:o,i[1]=c;for(var p=2;p<a;p++)i[p]=r[p];return n.createElement.apply(null,i)}return n.createElement.apply(null,r)}s.displayName="MDXCreateElement"},2508:function(e,t,r){r.r(t),r.d(t,{frontMatter:function(){return c},contentTitle:function(){return l},metadata:function(){return p},toc:function(){return u},default:function(){return s}});var n=r(7462),o=r(3366),a=(r(7294),r(3905)),i=["components"],c={},l=void 0,p={unversionedId:"overview-of-charting/06-06 ChartApplyFormattingToSelectedmd",id:"overview-of-charting/06-06 ChartApplyFormattingToSelectedmd",title:"06-06 ChartApplyFormattingToSelectedmd",description:"Chart_ApplyFormattingToSelected.md",source:"@site/docs/06-overview-of-charting/06-06 ChartApplyFormattingToSelectedmd.md",sourceDirName:"06-overview-of-charting",slug:"/overview-of-charting/06-06 ChartApplyFormattingToSelectedmd",permalink:"/docs/overview-of-charting/06-06 ChartApplyFormattingToSelectedmd",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/06-overview-of-charting/06-06 ChartApplyFormattingToSelectedmd.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"06-05 ChartAddTitlesmd",permalink:"/docs/overview-of-charting/06-05 ChartAddTitlesmd"},next:{title:"06-07 ChartApplyTrendColorsmd",permalink:"/docs/overview-of-charting/06-07 ChartApplyTrendColorsmd"}},u=[{value:"Chart_ApplyFormattingToSelected.md",id:"chart_applyformattingtoselectedmd",children:[],level:2}],d={toc:u};function s(e){var t=e.components,r=(0,o.Z)(e,i);return(0,a.kt)("wrapper",(0,n.Z)({},d,r,{components:t,mdxType:"MDXLayout"}),(0,a.kt)("h2",{id:"chart_applyformattingtoselectedmd"},"Chart_ApplyFormattingToSelected.md"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb"},"Public Sub Chart_ApplyFormattingToSelected()\n\n    Dim targetObject As ChartObject\n    Const MARKER_SIZE As Long = 5\n\n    For Each targetObject In Chart_GetObjectsFromObject(Selection)\n\n        Dim targetSeries As series\n\n        For Each targetSeries In targetObject.Chart.SeriesCollection\n            targetSeries.MarkerSize = MARKER_SIZE\n        Next targetSeries\n    Next targetObject\n\nEnd Sub\n")))}s.isMDXComponent=!0}}]);