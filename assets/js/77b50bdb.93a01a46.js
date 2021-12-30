"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[3589],{3905:function(e,t,r){r.d(t,{Zo:function(){return c},kt:function(){return h}});var n=r(7294);function a(e,t,r){return t in e?Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}):e[t]=r,e}function i(e,t){var r=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);t&&(n=n.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),r.push.apply(r,n)}return r}function s(e){for(var t=1;t<arguments.length;t++){var r=null!=arguments[t]?arguments[t]:{};t%2?i(Object(r),!0).forEach((function(t){a(e,t,r[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(r)):i(Object(r)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(r,t))}))}return e}function l(e,t){if(null==e)return{};var r,n,a=function(e,t){if(null==e)return{};var r,n,a={},i=Object.keys(e);for(n=0;n<i.length;n++)r=i[n],t.indexOf(r)>=0||(a[r]=e[r]);return a}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(n=0;n<i.length;n++)r=i[n],t.indexOf(r)>=0||Object.prototype.propertyIsEnumerable.call(e,r)&&(a[r]=e[r])}return a}var o=n.createContext({}),u=function(e){var t=n.useContext(o),r=t;return e&&(r="function"==typeof e?e(t):s(s({},t),e)),r},c=function(e){var t=u(e.components);return n.createElement(o.Provider,{value:t},e.children)},f={inlineCode:"code",wrapper:function(e){var t=e.children;return n.createElement(n.Fragment,{},t)}},m=n.forwardRef((function(e,t){var r=e.components,a=e.mdxType,i=e.originalType,o=e.parentName,c=l(e,["components","mdxType","originalType","parentName"]),m=u(r),h=a,p=m["".concat(o,".").concat(h)]||m[h]||f[h]||i;return r?n.createElement(p,s(s({ref:t},c),{},{components:r})):n.createElement(p,s({ref:t},c))}));function h(e,t){var r=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var i=r.length,s=new Array(i);s[0]=m;var l={};for(var o in t)hasOwnProperty.call(t,o)&&(l[o]=t[o]);l.originalType=e,l.mdxType="string"==typeof e?e:a,s[1]=l;for(var u=2;u<i;u++)s[u]=r[u];return n.createElement.apply(null,s)}return n.createElement.apply(null,r)}m.displayName="MDXCreateElement"},387:function(e,t,r){r.r(t),r.d(t,{frontMatter:function(){return l},contentTitle:function(){return o},metadata:function(){return u},toc:function(){return c},default:function(){return m}});var n=r(7462),a=r(3366),i=(r(7294),r(3905)),s=["components"],l={},o=void 0,u={unversionedId:"overview-of-charting/06-13 ChartFlipXYValuesmd",id:"overview-of-charting/06-13 ChartFlipXYValuesmd",title:"06-13 ChartFlipXYValuesmd",description:"ChartFlipXYValues.md",source:"@site/docs/06-overview-of-charting/06-13 ChartFlipXYValuesmd.md",sourceDirName:"06-overview-of-charting",slug:"/overview-of-charting/06-13 ChartFlipXYValuesmd",permalink:"/docs/overview-of-charting/06-13 ChartFlipXYValuesmd",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/06-overview-of-charting/06-13 ChartFlipXYValuesmd.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"06-12 ChartSortSeriesByNamemd",permalink:"/docs/overview-of-charting/06-12 ChartSortSeriesByNamemd"},next:{title:"06-14 ChartMergeSeriesmd",permalink:"/docs/overview-of-charting/06-14 ChartMergeSeriesmd"}},c=[{value:"ChartFlipXYValues.md",id:"chartflipxyvaluesmd",children:[],level:2}],f={toc:c};function m(e){var t=e.components,r=(0,a.Z)(e,s);return(0,i.kt)("wrapper",(0,n.Z)({},f,r,{components:t,mdxType:"MDXLayout"}),(0,i.kt)("h2",{id:"chartflipxyvaluesmd"},"ChartFlipXYValues.md"),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-vb"},"Public Sub ChartFlipXYValues()\n\n    Dim targetObject As ChartObject\n    Dim targetChart As Chart\n    For Each targetObject In Chart_GetObjectsFromObject(Selection)\n        Set targetChart = targetObject.Chart\n\n        Dim butlSeriesies As New Collection\n        Dim butlSeries As bUTLChartSeries\n\n        Dim targetSeries As series\n        For Each targetSeries In targetChart.SeriesCollection\n            Set butlSeries = New bUTLChartSeries\n            butlSeries.UpdateFromChartSeries targetSeries\n\n            Dim dummyRange As Range\n\n            Set dummyRange = butlSeries.Values\n            Set butlSeries.Values = butlSeries.XValues\n            Set butlSeries.XValues = dummyRange\n\n            'need to change the series name also\n            'assume that title is same offset\n            'code blocked for now\n            If False And Not butlSeries.name Is Nothing Then\n                Dim rowsOffset As Long, columnsOffset As Long\n                rowsOffset = butlSeries.name.Row - butlSeries.XValues.Cells(1, 1).Row\n                columnsOffset = butlSeries.name.Column - butlSeries.XValues.Cells(1, 1).Column\n\n                Set butlSeries.name = butlSeries.Values.Cells(1, 1).Offset(rowsOffset, columnsOffset)\n            End If\n\n            butlSeries.UpdateSeriesWithNewValues\n\n        Next targetSeries\n\n        ''need to flip axis labels if they exist\n        ''three cases: X only, Y only, X and Y\n\n        If targetChart.Axes(xlCategory).HasTitle And Not targetChart.Axes(xlValue).HasTitle Then\n\n            targetChart.Axes(xlValue).HasTitle = True\n            targetChart.Axes(xlValue).AxisTitle.Text = targetChart.Axes(xlCategory).AxisTitle.Text\n            targetChart.Axes(xlCategory).HasTitle = False\n\n        ElseIf Not targetChart.Axes(xlCategory).HasTitle And targetChart.Axes(xlValue).HasTitle Then\n            targetChart.Axes(xlCategory).HasTitle = True\n            targetChart.Axes(xlCategory).AxisTitle.Text = targetChart.Axes(xlValue).AxisTitle.Text\n            targetChart.Axes(xlValue).HasTitle = False\n\n        ElseIf targetChart.Axes(xlCategory).HasTitle And targetChart.Axes(xlValue).HasTitle Then\n            Dim swapText As String\n\n            swapText = targetChart.Axes(xlCategory).AxisTitle.Text\n\n            targetChart.Axes(xlCategory).AxisTitle.Text = targetChart.Axes(xlValue).AxisTitle.Text\n            targetChart.Axes(xlValue).AxisTitle.Text = swapText\n\n        End If\n\n        Set butlSeriesies = Nothing\n\n    Next targetObject\n\nEnd Sub\n")))}m.isMDXComponent=!0}}]);