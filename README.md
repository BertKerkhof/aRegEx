# aRegEx

Ergonomic improvement of regulated expression programming was a 
leading motive in the development of this module.

Best practice with RegEx patterns is to define beforehand the grand
patterns and its parts that are central to your programming mission.
Then let robotics make the modifications necessary for each RegEx
application. The new function LimitCap is made for this purpose. It
modifies a pattern for optimal employment.

The method to first collect multiple matches before results will be
processed, reduces complexity and is less error prone. The new
amRegExCapture function gives improved access to details concerning
multiple matches. For people familiar with modern arrays and list
processing, the returned result will be manageable. Note that the 
order of parameters in the functions mRegExdotHeader and 
mRegExdotFooter has changed.

With the practice to start programming with pattern design and the 
method to first collect multiple matches before processing, program
development time can be reduced.

A new function analyses the RegEx pattern: nCapture gives info before
RegEx execution. With the use of LimitCap and nCapture and numerous
small optimizations, execution speed of the module is shortened.

I hope the functions presented here will facilitate the use of RegEx
and make people curious about regulated expressions and what they can
do for you.

Author: kerkhof.bert@gmail.com
