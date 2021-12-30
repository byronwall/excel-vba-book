"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[4825],{3905:function(e,r,t){t.d(r,{Zo:function(){return u},kt:function(){return h}});var n=t(7294);function a(e,r,t){return r in e?Object.defineProperty(e,r,{value:t,enumerable:!0,configurable:!0,writable:!0}):e[r]=t,e}function i(e,r){var t=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);r&&(n=n.filter((function(r){return Object.getOwnPropertyDescriptor(e,r).enumerable}))),t.push.apply(t,n)}return t}function o(e){for(var r=1;r<arguments.length;r++){var t=null!=arguments[r]?arguments[r]:{};r%2?i(Object(t),!0).forEach((function(r){a(e,r,t[r])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(t)):i(Object(t)).forEach((function(r){Object.defineProperty(e,r,Object.getOwnPropertyDescriptor(t,r))}))}return e}function s(e,r){if(null==e)return{};var t,n,a=function(e,r){if(null==e)return{};var t,n,a={},i=Object.keys(e);for(n=0;n<i.length;n++)t=i[n],r.indexOf(t)>=0||(a[t]=e[t]);return a}(e,r);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(n=0;n<i.length;n++)t=i[n],r.indexOf(t)>=0||Object.prototype.propertyIsEnumerable.call(e,t)&&(a[t]=e[t])}return a}var c=n.createContext({}),l=function(e){var r=n.useContext(c),t=r;return e&&(t="function"==typeof e?e(r):o(o({},r),e)),t},u=function(e){var r=l(e.components);return n.createElement(c.Provider,{value:r},e.children)},p={inlineCode:"code",wrapper:function(e){var r=e.children;return n.createElement(n.Fragment,{},r)}},f=n.forwardRef((function(e,r){var t=e.components,a=e.mdxType,i=e.originalType,c=e.parentName,u=s(e,["components","mdxType","originalType","parentName"]),f=l(t),h=a,m=f["".concat(c,".").concat(h)]||f[h]||p[h]||i;return t?n.createElement(m,o(o({ref:r},u),{},{components:t})):n.createElement(m,o({ref:r},u))}));function h(e,r){var t=arguments,a=r&&r.mdxType;if("string"==typeof e||a){var i=t.length,o=new Array(i);o[0]=f;var s={};for(var c in r)hasOwnProperty.call(r,c)&&(s[c]=r[c]);s.originalType=e,s.mdxType="string"==typeof e?e:a,o[1]=s;for(var l=2;l<i;l++)o[l]=t[l];return n.createElement.apply(null,o)}return n.createElement.apply(null,t)}f.displayName="MDXCreateElement"},134:function(e,r,t){t.r(r),t.d(r,{frontMatter:function(){return s},contentTitle:function(){return c},metadata:function(){return l},toc:function(){return u},default:function(){return f}});var n=t(7462),a=t(3366),i=(t(7294),t(3905)),o=["components"],s={},c=void 0,l={unversionedId:"overview-of-charting/06-14 ChartMergeSeriesmd",id:"overview-of-charting/06-14 ChartMergeSeriesmd",title:"06-14 ChartMergeSeriesmd",description:"ChartMergeSeries.md",source:"@site/docs/06-overview-of-charting/06-14 ChartMergeSeriesmd.md",sourceDirName:"06-overview-of-charting",slug:"/overview-of-charting/06-14 ChartMergeSeriesmd",permalink:"/docs/overview-of-charting/06-14 ChartMergeSeriesmd",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/06-overview-of-charting/06-14 ChartMergeSeriesmd.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"06-13 ChartFlipXYValuesmd",permalink:"/docs/overview-of-charting/06-13 ChartFlipXYValuesmd"},next:{title:"06-15 ChartSplitSeriesmd",permalink:"/docs/overview-of-charting/06-15 ChartSplitSeriesmd"}},u=[{value:"ChartMergeSeries.md",id:"chartmergeseriesmd",children:[],level:2}],p={toc:u};function f(e){var r=e.components,t=(0,a.Z)(e,o);return(0,i.kt)("wrapper",(0,n.Z)({},p,t,{components:r,mdxType:"MDXLayout"}),(0,i.kt)("h2",{id:"chartmergeseriesmd"},"ChartMergeSeries.md"),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-vb"},"Public Sub ChartMergeSeries()\n\n    Dim targetObject As ChartObject\n    Dim targetChart As Chart\n    Dim firstChart As Chart\n\n    Dim isFirstChart As Boolean\n    isFirstChart = True\n\n    Application.ScreenUpdating = False\n\n    For Each targetObject In Chart_GetObjectsFromObject(Selection)\n\n        Set targetChart = targetObject.Chart\n        If isFirstChart Then\n            Set firstChart = targetChart\n            isFirstChart = False\n        Else\n            Dim targetSeries As series\n            For Each targetSeries In targetChart.SeriesCollection\n\n                Dim newChartSeries As series\n                Dim butlSeries As New bUTLChartSeries\n\n                butlSeries.UpdateFromChartSeries targetSeries\n                Set newChartSeries = butlSeries.AddSeriesToChart(firstChart)\n\n                newChartSeries.MarkerSize = targetSeries.MarkerSize\n                newChartSeries.MarkerStyle = targetSeries.MarkerStyle\n\n                targetSeries.Delete\n\n            Next targetSeries\n\n            targetObject.Delete\n\n        End If\n    Next targetObject\n\n    Application.ScreenUpdating = True\n\nEnd Sub\n")))}f.isMDXComponent=!0}}]);