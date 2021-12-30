"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[18],{3905:function(e,r,n){n.d(r,{Zo:function(){return s},kt:function(){return f}});var t=n(7294);function o(e,r,n){return r in e?Object.defineProperty(e,r,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[r]=n,e}function a(e,r){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var t=Object.getOwnPropertySymbols(e);r&&(t=t.filter((function(r){return Object.getOwnPropertyDescriptor(e,r).enumerable}))),n.push.apply(n,t)}return n}function l(e){for(var r=1;r<arguments.length;r++){var n=null!=arguments[r]?arguments[r]:{};r%2?a(Object(n),!0).forEach((function(r){o(e,r,n[r])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):a(Object(n)).forEach((function(r){Object.defineProperty(e,r,Object.getOwnPropertyDescriptor(n,r))}))}return e}function i(e,r){if(null==e)return{};var n,t,o=function(e,r){if(null==e)return{};var n,t,o={},a=Object.keys(e);for(t=0;t<a.length;t++)n=a[t],r.indexOf(n)>=0||(o[n]=e[n]);return o}(e,r);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(t=0;t<a.length;t++)n=a[t],r.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(o[n]=e[n])}return o}var u=t.createContext({}),c=function(e){var r=t.useContext(u),n=r;return e&&(n="function"==typeof e?e(r):l(l({},r),e)),n},s=function(e){var r=c(e.components);return t.createElement(u.Provider,{value:r},e.children)},p={inlineCode:"code",wrapper:function(e){var r=e.children;return t.createElement(t.Fragment,{},r)}},d=t.forwardRef((function(e,r){var n=e.components,o=e.mdxType,a=e.originalType,u=e.parentName,s=i(e,["components","mdxType","originalType","parentName"]),d=c(n),f=o,m=d["".concat(u,".").concat(f)]||d[f]||p[f]||a;return n?t.createElement(m,l(l({ref:r},s),{},{components:n})):t.createElement(m,l({ref:r},s))}));function f(e,r){var n=arguments,o=r&&r.mdxType;if("string"==typeof e||o){var a=n.length,l=new Array(a);l[0]=d;var i={};for(var u in r)hasOwnProperty.call(r,u)&&(i[u]=r[u]);i.originalType=e,i.mdxType="string"==typeof e?e:o,l[1]=i;for(var c=2;c<a;c++)l[c]=n[c];return t.createElement.apply(null,l)}return t.createElement.apply(null,n)}d.displayName="MDXCreateElement"},9198:function(e,r,n){n.r(r),n.d(r,{frontMatter:function(){return i},contentTitle:function(){return u},metadata:function(){return c},toc:function(){return s},default:function(){return d}});var t=n(7462),o=n(3366),a=(n(7294),n(3905)),l=["components"],i={},u=void 0,c={unversionedId:"overview-of-values-and-formulas/05-13 MakeHyperlinksmd",id:"overview-of-values-and-formulas/05-13 MakeHyperlinksmd",title:"05-13 MakeHyperlinksmd",description:"MakeHyperlinks.md",source:"@site/docs/05-overview-of-values-and-formulas/05-13 MakeHyperlinksmd.md",sourceDirName:"05-overview-of-values-and-formulas",slug:"/overview-of-values-and-formulas/05-13 MakeHyperlinksmd",permalink:"/docs/overview-of-values-and-formulas/05-13 MakeHyperlinksmd",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/05-overview-of-values-and-formulas/05-13 MakeHyperlinksmd.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"05-12 SplitIntoColumnsmd",permalink:"/docs/overview-of-values-and-formulas/05-12 SplitIntoColumnsmd"},next:{title:"05-13 SplitIntoRowsmd",permalink:"/docs/overview-of-values-and-formulas/05-13 SplitIntoRowsmd"}},s=[{value:"MakeHyperlinks.md",id:"makehyperlinksmd",children:[],level:2}],p={toc:s};function d(e){var r=e.components,n=(0,o.Z)(e,l);return(0,a.kt)("wrapper",(0,t.Z)({},p,n,{components:r,mdxType:"MDXLayout"}),(0,a.kt)("h2",{id:"makehyperlinksmd"},"MakeHyperlinks.md"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb"},'Public Sub MakeHyperlinks()\n\n    \'+Changed to inputbox\n    On Error GoTo errHandler\n    Dim targetRange As Range\n    Set targetRange = GetInputOrSelection("Select the range of cells to convert to hyperlink")\n\n    \'TODO: choose a better variable name\n    Dim targetCell As Range\n    For Each targetCell In targetRange\n        ActiveSheet.Hyperlinks.Add Anchor:=targetCell, Address:=targetCell\n    Next targetCell\n    Exit Sub\nerrHandler:\n    MsgBox "No Range Selected!"\nEnd Sub\n')))}d.isMDXComponent=!0}}]);