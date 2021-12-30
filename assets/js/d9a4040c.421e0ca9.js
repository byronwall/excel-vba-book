"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[9663],{3905:function(e,t,n){n.d(t,{Zo:function(){return u},kt:function(){return f}});var r=n(7294);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function a(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);t&&(r=r.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,r)}return n}function i(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?a(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):a(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function s(e,t){if(null==e)return{};var n,r,o=function(e,t){if(null==e)return{};var n,r,o={},a=Object.keys(e);for(r=0;r<a.length;r++)n=a[r],t.indexOf(n)>=0||(o[n]=e[n]);return o}(e,t);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(r=0;r<a.length;r++)n=a[r],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(o[n]=e[n])}return o}var c=r.createContext({}),l=function(e){var t=r.useContext(c),n=t;return e&&(n="function"==typeof e?e(t):i(i({},t),e)),n},u=function(e){var t=l(e.components);return r.createElement(c.Provider,{value:t},e.children)},p={inlineCode:"code",wrapper:function(e){var t=e.children;return r.createElement(r.Fragment,{},t)}},d=r.forwardRef((function(e,t){var n=e.components,o=e.mdxType,a=e.originalType,c=e.parentName,u=s(e,["components","mdxType","originalType","parentName"]),d=l(n),f=o,v=d["".concat(c,".").concat(f)]||d[f]||p[f]||a;return n?r.createElement(v,i(i({ref:t},u),{},{components:n})):r.createElement(v,i({ref:t},u))}));function f(e,t){var n=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var a=n.length,i=new Array(a);i[0]=d;var s={};for(var c in t)hasOwnProperty.call(t,c)&&(s[c]=t[c]);s.originalType=e,s.mdxType="string"==typeof e?e:o,i[1]=s;for(var l=2;l<a;l++)i[l]=n[l];return r.createElement.apply(null,i)}return r.createElement.apply(null,n)}d.displayName="MDXCreateElement"},5262:function(e,t,n){n.r(t),n.d(t,{frontMatter:function(){return s},contentTitle:function(){return c},metadata:function(){return l},toc:function(){return u},default:function(){return d}});var r=n(7462),o=n(3366),a=(n(7294),n(3905)),i=["components"],s={},c=void 0,l={unversionedId:"overview-of-events/11-03 common-patterns",id:"overview-of-events/11-03 common-patterns",title:"11-03 common-patterns",description:"common patterns",source:"@site/docs/11-overview-of-events/11-03 common-patterns.md",sourceDirName:"11-overview-of-events",slug:"/overview-of-events/11-03 common-patterns",permalink:"/docs/overview-of-events/11-03 common-patterns",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/11-overview-of-events/11-03 common-patterns.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"11-02 specific-events",permalink:"/docs/overview-of-events/11-02 specific-events"},next:{title:"11-04 more-advanced-events",permalink:"/docs/overview-of-events/11-04 more-advanced-events"}},u=[{value:"common patterns",id:"common-patterns",children:[{value:"Intersect",id:"intersect",children:[],level:3},{value:"Application.EnableEvents = False",id:"applicationenableevents--false",children:[],level:3}],level:2}],p={toc:u};function d(e){var t=e.components,n=(0,o.Z)(e,i);return(0,a.kt)("wrapper",(0,r.Z)({},p,n,{components:t,mdxType:"MDXLayout"}),(0,a.kt)("h2",{id:"common-patterns"},"common patterns"),(0,a.kt)("p",null,"There are a number of patterns that are very common with Events. These patterns typically exist to avoid causing a problem or to avoid extra work where possible. Most VBA is not performance critical, but it is possible for an event to be called hundreds of times for a given chucnk of code. Since this is true, you can start to have an immediate impact on performance if your event handling code includes a number of unnecessary steps. As a side note, this is a good reminder that when trying to speed up code, you will nearly always do better to add ",(0,a.kt)("inlineCode",{parentName:"p"},"Application.EnableEvents = False")," before your performance critical code; this assumes that your VBA does not rely on events firing to function properly."),(0,a.kt)("h3",{id:"intersect"},"Intersect"),(0,a.kt)("p",null,"The first is the ",(0,a.kt)("inlineCode",{parentName:"p"},"Intersect"),' technique to determine if a Range that was affected by an event was a Range of interest. With this approach, you define a Range which includes your "interesting" cells. You then do a ',(0,a.kt)("inlineCode",{parentName:"p"},"If Not Intersect(rngEvent, rngTarget) Is Nothing")," to see if the intersection of the callback Range and the desired Range overlap. If they overlap, yhen you typically execute some code. This allows you to quickly filter out Ranges which have changed but are not relevant to Athena code you need to run."),(0,a.kt)("p",null,"TODO: add a code sample here"),(0,a.kt)("h3",{id:"applicationenableevents--false"},"Application.EnableEvents = False"),(0,a.kt)("p",null,"One of the biggest gotchas with Events is that you can quickly and accidentally create an endless loop of Event code running if your event handler is able to retirgger the original event. This is quite common if you are looking at the Selection and then change the selected cell. The same can happen if you are using an event to watch for a change and then you respond with additional changes. Both of these accidents are so common, that you should seriously consider always disabling events in your handler. It is quite rare that you will need an other event to fire following your own processing."),(0,a.kt)("p",null,"The main thing to remember here is that you really need to enable events again. Excel will not do this for you. You can create odd situations if you have an error in your code that goes unchecked. This situation can mean that events are disabled. For really sensitive, user focused code, you should add a proper error handler and enable events following that."),(0,a.kt)("p",null,"To handle this event, the code is quite simple:"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb"},"Sub EventHandler()\n    'disable events\n    Application.EnableEvents = False\n\n    '' do some stuff\n\n    're-enable events\n    Application.EnableEvents = True\nEnd Sub\n")))}d.isMDXComponent=!0}}]);