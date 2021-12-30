"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[5512],{3905:function(e,r,t){t.d(r,{Zo:function(){return l},kt:function(){return f}});var n=t(7294);function o(e,r,t){return r in e?Object.defineProperty(e,r,{value:t,enumerable:!0,configurable:!0,writable:!0}):e[r]=t,e}function a(e,r){var t=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);r&&(n=n.filter((function(r){return Object.getOwnPropertyDescriptor(e,r).enumerable}))),t.push.apply(t,n)}return t}function i(e){for(var r=1;r<arguments.length;r++){var t=null!=arguments[r]?arguments[r]:{};r%2?a(Object(t),!0).forEach((function(r){o(e,r,t[r])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(t)):a(Object(t)).forEach((function(r){Object.defineProperty(e,r,Object.getOwnPropertyDescriptor(t,r))}))}return e}function u(e,r){if(null==e)return{};var t,n,o=function(e,r){if(null==e)return{};var t,n,o={},a=Object.keys(e);for(n=0;n<a.length;n++)t=a[n],r.indexOf(t)>=0||(o[t]=e[t]);return o}(e,r);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(n=0;n<a.length;n++)t=a[n],r.indexOf(t)>=0||Object.prototype.propertyIsEnumerable.call(e,t)&&(o[t]=e[t])}return o}var s=n.createContext({}),c=function(e){var r=n.useContext(s),t=r;return e&&(t="function"==typeof e?e(r):i(i({},r),e)),t},l=function(e){var r=c(e.components);return n.createElement(s.Provider,{value:r},e.children)},m={inlineCode:"code",wrapper:function(e){var r=e.children;return n.createElement(n.Fragment,{},r)}},p=n.forwardRef((function(e,r){var t=e.components,o=e.mdxType,a=e.originalType,s=e.parentName,l=u(e,["components","mdxType","originalType","parentName"]),p=c(t),f=o,h=p["".concat(s,".").concat(f)]||p[f]||m[f]||a;return t?n.createElement(h,i(i({ref:r},l),{},{components:t})):n.createElement(h,i({ref:r},l))}));function f(e,r){var t=arguments,o=r&&r.mdxType;if("string"==typeof e||o){var a=t.length,i=new Array(a);i[0]=p;var u={};for(var s in r)hasOwnProperty.call(r,s)&&(u[s]=r[s]);u.originalType=e,u.mdxType="string"==typeof e?e:o,i[1]=u;for(var c=2;c<a;c++)i[c]=t[c];return n.createElement.apply(null,i)}return n.createElement.apply(null,t)}p.displayName="MDXCreateElement"},7694:function(e,r,t){t.r(r),t.d(r,{frontMatter:function(){return u},contentTitle:function(){return s},metadata:function(){return c},toc:function(){return l},default:function(){return p}});var n=t(7462),o=t(3366),a=(t(7294),t(3905)),i=["components"],u={},s=void 0,c={unversionedId:"overview-of-user-forms-and-input/12-03 making-that-USerForm-show-up",id:"overview-of-user-forms-and-input/12-03 making-that-USerForm-show-up",title:"12-03 making-that-USerForm-show-up",description:"making that USerForm show up",source:"@site/docs/12-overview-of-user-forms-and-input/12-03 making-that-USerForm-show-up.md",sourceDirName:"12-overview-of-user-forms-and-input",slug:"/overview-of-user-forms-and-input/12-03 making-that-USerForm-show-up",permalink:"/excel-vba-book/docs/overview-of-user-forms-and-input/12-03 making-that-USerForm-show-up",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/12-overview-of-user-forms-and-input/12-03 making-that-USerForm-show-up.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"12-02 creating-a-UserForm",permalink:"/excel-vba-book/docs/overview-of-user-forms-and-input/12-02 creating-a-UserForm"},next:{title:"12-04 adding-controls-to-a-UserForm-and-wiring-them-up",permalink:"/excel-vba-book/docs/overview-of-user-forms-and-input/12-04 adding-controls-to-a-UserForm-and-wiring-them-up"}},l=[{value:"making that USerForm show up",id:"making-that-userform-show-up",children:[],level:2}],m={toc:l};function p(e){var r=e.components,t=(0,o.Z)(e,i);return(0,a.kt)("wrapper",(0,n.Z)({},m,t,{components:r,mdxType:"MDXLayout"}),(0,a.kt)("h2",{id:"making-that-userform-show-up"},"making that USerForm show up"),(0,a.kt)("p",null,"Once your UserForm is created, there are a couple of ways of showing it on screen:"),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},"Run any code from the VBE that is contained within the form. This will show the form."),(0,a.kt)("li",{parentName:"ul"},"Create an instance of the form somewhere and show it")),(0,a.kt)("p",null,'For those two methods, the latter is really the only one that will work for user applications or other "real" uses. If you are simply testing or doing things for yourself, then hitting F5 in the VBE may not be a large ask.'),(0,a.kt)("p",null,"For the former, see the code below for an example of how to show the form."),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb"},"DIm frm as UserForm\nSet frm = New UserForm\n\nfrm.Show\n")))}p.isMDXComponent=!0}}]);