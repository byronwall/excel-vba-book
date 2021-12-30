"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[2733],{3905:function(e,n,t){t.d(n,{Zo:function(){return c},kt:function(){return d}});var r=t(7294);function o(e,n,t){return n in e?Object.defineProperty(e,n,{value:t,enumerable:!0,configurable:!0,writable:!0}):e[n]=t,e}function i(e,n){var t=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);n&&(r=r.filter((function(n){return Object.getOwnPropertyDescriptor(e,n).enumerable}))),t.push.apply(t,r)}return t}function a(e){for(var n=1;n<arguments.length;n++){var t=null!=arguments[n]?arguments[n]:{};n%2?i(Object(t),!0).forEach((function(n){o(e,n,t[n])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(t)):i(Object(t)).forEach((function(n){Object.defineProperty(e,n,Object.getOwnPropertyDescriptor(t,n))}))}return e}function l(e,n){if(null==e)return{};var t,r,o=function(e,n){if(null==e)return{};var t,r,o={},i=Object.keys(e);for(r=0;r<i.length;r++)t=i[r],n.indexOf(t)>=0||(o[t]=e[t]);return o}(e,n);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(r=0;r<i.length;r++)t=i[r],n.indexOf(t)>=0||Object.prototype.propertyIsEnumerable.call(e,t)&&(o[t]=e[t])}return o}var u=r.createContext({}),s=function(e){var n=r.useContext(u),t=n;return e&&(t="function"==typeof e?e(n):a(a({},n),e)),t},c=function(e){var n=s(e.components);return r.createElement(u.Provider,{value:n},e.children)},m={inlineCode:"code",wrapper:function(e){var n=e.children;return r.createElement(r.Fragment,{},n)}},p=r.forwardRef((function(e,n){var t=e.components,o=e.mdxType,i=e.originalType,u=e.parentName,c=l(e,["components","mdxType","originalType","parentName"]),p=s(t),d=o,f=p["".concat(u,".").concat(d)]||p[d]||m[d]||i;return t?r.createElement(f,a(a({ref:n},c),{},{components:t})):r.createElement(f,a({ref:n},c))}));function d(e,n){var t=arguments,o=n&&n.mdxType;if("string"==typeof e||o){var i=t.length,a=new Array(i);a[0]=p;var l={};for(var u in n)hasOwnProperty.call(n,u)&&(l[u]=n[u]);l.originalType=e,l.mdxType="string"==typeof e?e:o,a[1]=l;for(var s=2;s<i;s++)a[s]=t[s];return r.createElement.apply(null,a)}return r.createElement.apply(null,t)}p.displayName="MDXCreateElement"},9906:function(e,n,t){t.r(n),t.d(n,{frontMatter:function(){return l},contentTitle:function(){return u},metadata:function(){return s},toc:function(){return c},default:function(){return p}});var r=t(7462),o=t(3366),i=(t(7294),t(3905)),a=["components"],l={},u=void 0,s={unversionedId:"overview-of-utility-code/15-12 SeriesSplitIntoBinsmd",id:"overview-of-utility-code/15-12 SeriesSplitIntoBinsmd",title:"15-12 SeriesSplitIntoBinsmd",description:"SeriesSplitIntoBins.md",source:"@site/docs/15-overview-of-utility-code/15-12 SeriesSplitIntoBinsmd.md",sourceDirName:"15-overview-of-utility-code",slug:"/overview-of-utility-code/15-12 SeriesSplitIntoBinsmd",permalink:"/docs/overview-of-utility-code/15-12 SeriesSplitIntoBinsmd",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/15-overview-of-utility-code/15-12 SeriesSplitIntoBinsmd.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"15-11 SeriesSplitmd",permalink:"/docs/overview-of-utility-code/15-11 SeriesSplitmd"},next:{title:"15-13 SheetDeleteHiddenRowsmd",permalink:"/docs/overview-of-utility-code/15-13 SheetDeleteHiddenRowsmd"}},c=[{value:"SeriesSplitIntoBins.md",id:"seriessplitintobinsmd",children:[],level:2}],m={toc:c};function p(e){var n=e.components,t=(0,o.Z)(e,a);return(0,i.kt)("wrapper",(0,r.Z)({},m,t,{components:n,mdxType:"MDXLayout"}),(0,i.kt)("h2",{id:"seriessplitintobinsmd"},"SeriesSplitIntoBins.md"),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-vb"},'Public Sub SeriesSplitIntoBins()\n\n    Const LESS_THAN_EQUAL_TO_GENERAL As String = "<= General"\n    Const GREATER_THAN_GENERAL As String = "> General"\n    On Error GoTo ErrorNoSelection\n\n    Dim selectedRange As Range\n    Set selectedRange = Application.InputBox("Select category range with heading", Type:=8)\n    Set selectedRange = Intersect(selectedRange, selectedRange.Parent.UsedRange) _\n                                 .SpecialCells(xlCellTypeVisible, xlLogical + _\n                                  xlNumbers + xlTextValues)\n\n    Dim valueRange As Range\n    Set valueRange = Application.InputBox("Select values range with heading", Type:=8)\n    Set valueRange = Intersect(valueRange, valueRange.Parent.UsedRange)\n\n    \'\'need to prompt for max/min/bins\n    Dim maximumValue As Double, minimumValue As Double, binValue As Long\n\n    minimumValue = Application.InputBox("Minimum value.", "Min", _\n                                        WorksheetFunction.Min(selectedRange), Type:=1)\n\n    maximumValue = Application.InputBox("Maximum value.", "Max", _\n                                        WorksheetFunction.Max(selectedRange), Type:=1)\n\n    binValue = Application.InputBox("Number of groups.", "Bins", _\n                                    WorksheetFunction.RoundDown(Math.Sqr(WorksheetFunction.Count(selectedRange)), _\n                                    0), Type:=1)\n\n    On Error GoTo 0\n\n    \'determine default value\n    Dim defaultString As Variant\n    defaultString = Application.InputBox("Enter the default value", "Default", "#N/A")\n\n    \'detect cancel and exit\n    If StrPtr(defaultString) = 0 Then Exit Sub\n\n    \'\'TODO prompt for output location\n\n    valueRange.EntireColumn.Offset(, 1).Resize(, binValue + 2).Insert\n    \'head the columns with the values\n\n    \'\'TODO add a For loop to go through the bins\n\n    Dim targetBin As Long\n    For targetBin = 0 To binValue\n        valueRange.Cells(1).Offset(, targetBin + 1) = minimumValue + (maximumValue - _\n                                                      minimumValue) * targetBin / binValue\n    Next\n\n    \'add the last item\n    valueRange.Cells(1).Offset(, binValue + 2).FormulaR1C1 = "=RC[-1]"\n\n    \'FIRST =IF($D2 <=V$1,$U2,#N/A)\n    \'=IF(RC4 <=R1C,RC21,#N/A)\n\n    \'MID =IF(AND($D2 <=W$1, $D2>V$1),$U2,#N/A)  \'\'\'W current, then left\n    \'=IF(AND(RC4 <=R1C, RC4>R1C[-1]),RC21,#N/A)\n\n    \'LAST =IF($D2>AA$1,$U2,#N/A)\n    \'=IF(RC4>R1C[-1],RC21,#N/A)\n\n    \'\'TODO add number format to display header correctly (helps with charts)\n\n    \'put the formula in for each column\n    \'=IF(RC13=R1C,RC16,#N/A)\n    Dim formulaHolder As Variant\n    formulaHolder = "=IF(AND(RC" & selectedRange.Column & " <=R" & _\n                    valueRange.Cells(1).Row & "C," & "RC" & selectedRange.Column & ">R" & _\n                    valueRange.Cells(1).Row & "C[-1]" & ")" & ",RC" & valueRange.Column & "," & _\n                    defaultString & ")"\n\n    Dim firstFormula As Variant\n    firstFormula = "=IF(AND(RC" & selectedRange.Column & " <=R" & _\n                    valueRange.Cells(1).Row & "C)" & ",RC" & valueRange.Column & "," & defaultString _\n                    & ")"\n\n    Dim lastFormula As Variant\n    lastFormula = "=IF(AND(RC" & selectedRange.Column & " >R" & _\n                    valueRange.Cells(1).Row & "C)" & ",RC" & valueRange.Column & "," & defaultString _\n                    & ")"\n\n    Dim formulaRange As Range\n    Set formulaRange = valueRange.Offset(1, 1).Resize(valueRange.Rows.Count - 1, binValue + 2)\n    formulaRange.FormulaR1C1 = formulaHolder\n\n    \'override with first/last\n    formulaRange.Columns(1).FormulaR1C1 = firstFormula\n    formulaRange.Columns(formulaRange.Columns.Count).FormulaR1C1 = lastFormula\n\n    formulaRange.EntireColumn.AutoFit\n\n    \'set the number formats\n\n    formulaRange.Offset(-1).Rows(1).Resize(1, binValue + 1).NumberFormat = LESS_THAN_EQUAL_TO_GENERAL\n    formulaRange.Offset(-1).Rows(1).Offset(, binValue + 1).NumberFormat = GREATER_THAN_GENERAL\n\n    Exit Sub\n\nErrorNoSelection:\n    \'TODO: consider removing this prompt\n    MsgBox "No selection made.  Exiting.", , "No selection"\n\nEnd Sub\n')))}p.isMDXComponent=!0}}]);