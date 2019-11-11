#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=aRegEx.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=Regulated expression wrapper
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Bert Kerkhof 2019-11-11 Apache 2.0 license
#AutoIt3Wrapper_Res_SaveSource=n
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 3 -w 4 -w 5
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/tc 2 /reel
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/pe /sf /sv /rm
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; The latest version of this source is published at GitHub in the
; BertKerkhof repository "aRegEx".

#include-once
#include <StringConstants.au3> ; Delivered with AutoIT
#include <aDrosteArray.au3>; Published at GitHub
#include <StringSupport.au3>; Published at GitHub

; Author ....: Bert Kerkhof ( kerkhof.bert@gmail.com )
; Tested with: AutoIt version 3.3.14.5 and win10

;
; Regulated expression interface with procedural syntax =============

; Advanced search
; As more and more digital texts pile up, the need to find a needle in a
; haystack increases. With regulated expressions (RegEx) one can effectively
; do an advanced search in large amounts of text files and find that needle.
;
; Structure informal data
; Lots of more or less structured data appear. Some people see a pattern and
; may want to get things in order. RegEx can transform stuff, bring things
; in a format suitable for further computer processing.
;
; Reformat between applications
; Many times data needs a transformation into a format suitable for another
; application. RegEx can do just that.
;
; Speed
; The RegEx technique works with interpreter programming packages as well as
; with compilers. Because a RegEx module is a kind of compiler in itself,
; it will do tiny character byte operations in nano-seconds where procedural
; syntax of interpreter programming is much slower. That is why RegEx is not
; to be missed.
;
; Career
; Computer programming languages differ in many respects. Programmers that
; learn to use RegEx will see that the syntax and the workings are much
; alike on many platforms.

; Standardization
; Instant data formats were created out of necessity, without much
; discussion. For example movie subtitle formats such as .srt .sub and .ass
; did see the light, for programmers and users it is a maze what exactly to
; expect. RegEx pattern script can describe the formats precisely. A formal
; description will make data more interchangeable between applications.
; Please communicate your experiences with patterns.
;
; aRegEx.au3
; This source offers several functions to work with RegEx in AutoIT. They
; all have 'RegEx in their name. One can derive the value type they return
; from the first letter:

;     i  denotes an index number pointing to the input string
;     s  string
;     n  count of the number of matches
;     p  RegEx pattern
;     a  array
;     m  array of RegEx captures and RegEx properties
;     am array of mRegEx array's
;
; The first four functions in the list below search for the first match. The
; next ones can find multiple matches. Many functions have the same input
; parameters, which eases the use:
;
; Function             Input-parameters                   Return value
; ---------------------------------------------------------------------
; To find a single match:
;   mRegEx()           $sIn $Pattern [$iStart]            $mArray
;   iRegEx()           $sIn $Pattern [$iStart]            $iIndex
;   sRegEx()           $sIn $Pattern [$iStart]  [$sSep]   $sString
;
; To operate multiple matches:
;   nRegEx()           $sIn $Pattern [$iStart]            $nCount
;   amRegExCapture()   $sIn $Pattern [$iStart]            $amCapture
;   aRegExLocate()     $sIn $Pattern [$iStart]  [$Lfirst] $aLocations
;   aRegExSplit()      $sIn $Pattern [$iStart]            $sParts
;   sRegExReplace()    $sIn $Pattern [$sScript] [$iStart] $sResult
;   aRegExCapture()    $sIn $Pattern [$iStart]            $a/aaCapture
;
; To operate on a RegEx pattern:
;   nCapture()         $Pattern                           $pSimple
;   LimitCap()         $Pattern [$iKeep1]..[$iKeep3]      $pBasic
;
; For use together with mRegEx and amRegExCapture():
;   mRegExdotHeader()  $mArray $sIn                       $sHeader
;   mRegExdotFooter()  $mArray $sIn                       $sFooter
;   mRegExdotContent() $mArray [$sSep]                    $sContent
;   mRegExdotFound()   $mArray                            $iStart
;   mRegExdotLast()    $mArray                            $iLast
;   mRegExdotLength()  $mArray                            $Length

; ---------------------------------------------------------------------

; The first input parameters $sInput and $Pattern are the same for many
; functions. From the third parameter on, the parameters are optional.
;
; For compatibility with an earlier implementation, some of the functions
; return a value via snapchat variable @extended:
;    + With sRegEx(), @extended return the start location of the match in
;      the input string.
;    + In iRegEx(), @extended is used as a pointer to the location
;      directly after the last match.
;    + With sRegExReplace(), the number of replacements made is delivered
;      in @extended.
; This use of @extended is deprecated. In future versions this
; functionality may cease.
;
; To help learn RegEx patterns, these Regex functions issue an informative
; halt message to the console if a pattern has a syntax error. @error code
; 2 is immediately intercepted within the CheckPattern function and
; brought to the attention of the programmer via the Scite console.
; Patterns with a zero length match may cause indefinite searches. Also
; these patterns are early intercepted. So the programmer never has to
; bother about further processing that would obscure the cause. This is a
; time saver.
;
; All the functions in this module never raise the @error flag also if
; there is no match. Some functions don't return matches at all. With
; other functions, the mFOUND property signals whether there is a match.
; For functions that process multiple matches, see the count property of
; the returned array.
;
; Patterns
; To practice with the formation of patterns, there is a pattern test
; program in my RegExGui repository at GitHub. Also the AutoIT package
; offers a helper GUI. Find it by searching on: "StringRegEpxGUI in the
; help files, then browse the page to "Resources" under "Example 6".
;
; Excellent documentation is at: "http://regular-expressions.info"
; and in: "Regular Expressions Cookbook" written by Jan Goyvaerts and Steven
; Levithan. The hard-copy is published by O'Reilly and available via Amazon.
;
; Unicode
; All the functions in this interface operate on Unicode utf input strings
; by default. To comply, the first thing to do is: save code with your
; source code editor in utf8 format. For old-style ascii input, precede the
; pattern with syntax: (*ASC)

;
; Version1 to version2 change log =====================================
; Ergonomic improvement of regulated expression programming was a
; leading motive.
;
; Best practice with RegEx patterns is to define beforehand the grand
; patterns and its parts that are central to your programming mission.
; Then let robotics make the modifications necessary for each RegEx
; application. The new function LimitCap is made for this purpose. It
; modifies a pattern for optimal employment.

; The method to first collect multiple matches before results will be
; processed, reduces complexity and is less error prone. The new
; amRegExCapture function gives improved access to details concerning
; multiple matches. For people familiar with modern arrays and list
; processing, the returned result will be manageable. Note that the
; order of parameters in the functions mRegExdotHeader and
; mRegExdotFooter has changed.
;
; With the practice to start programming with pattern design and the
; method to first collect multiple matches before processing, program
; development time can be reduced.
;
; A new function analyses the RegEx pattern: nCapture gives info before
; RegEx execution. With the use of LimitCap and nCapture and numerous
; small optimizations, the execution speed of the module increased
; considerable.
;
; I hope the functions presented here will facilitate the use of RegEx
; and make people curious about regulated expressions and what they can
; do for you.
;
; Author: kerkhof.bert@gmail.com

;
; Functions that operate on a RegEx pattern =========================

; Note: The nCapture and LimitCap functions are designed to operate
; with a fixed number of capture fields per pattern. This is frequently
; the case in applications based on the RegEx module delivered with
; AutoIT package v3, PCRE version Januari 2014.
;
; If using a pattern with a conditional numbered capture field such
; as (abc)? or (abc){0,1}, the nCapture and LimitCap functions are not
; accurate. Anyway the usage of such a conditional field is not
; advisable since it is difficult to find back the resulting captured
; data in the outputted array. Use ((?:abc)?) instead.

; #FUNCTION#
; Name ..........: nCapture
; Description ...: Calculate number of captures from the RegEx pattern
; Syntax ........: nCapture($Pattern)
; Parameters ....: $Pattern... RegEx pattern
; Returns .......: Number of captures
; Author ........: Bert Kerkhof
; Comment .......: If the pattern contains no capture field, the full
;                  match counts as one.

Func nCapture($Pattern)
  Local $N = UBound(StringRegExp($Pattern, "\((?![\?\*])", $STR_REGEXPARRAYGLOBALMATCH))
  Return $N ? $N : 1
EndFunc   ;==>nCapture

; #FUNCTION#
; Name ..........: LimitCap
; Description ...: Changes the pattern to receive less captures per match
; Syntax ........: LimitCap($Pattern[, $iKeep1 = 0][, $iKeep2 = 0][, $iKeep3 = 0])
; Parameters ....: $Pattern .. RegEx pattern.
;                  $iKeep1..3  [optional] numbers of a capture field to retain.
;                              Default is all 0 (keep no capture field).
; Returns .......: The modified pattern
; Author ........: Bert Kerkhof

Func LimitCap($Pattern, Const $iKeep1 = 0, Const $iKeep2 = 0, Const $iKeep3 = 0)
  Local $aKeep = aLeft(Array($iKeep1, $iKeep2, $iKeep3), @NumParams - 1)
  Local $K = 0, $aHook = aNew()
  While True
    $K = iRegEx($Pattern, "(?<!\\)\((?=[^\?\*])", $K + 1)
    If $K = 0 Then ExitLoop
    aAdd($aHook, $K)
  WEnd
  For $I = $aHook[0] To 1 Step -1
    If aSearch($aKeep, $I) Then ContinueLoop
    $Pattern = StringLeft($Pattern, $aHook[$I] - 1) & "(?:" & StringMid($Pattern, $aHook[$I] + 1)
  Next
  Return $Pattern
EndFunc   ;==>LimitCap

; #FUNCTION#
; Name ..........: CheckPattern
; Description ...: Checks whether the pattern is valid
; Syntax ........: CheckPattern($Pattern)
; Parameter .....: $Pattern ... RegEx pattern.
; Returns .......: None
; Author ........: Bert Kerkhof
; Comment .......: In case of an error, the program ends

Func CheckPattern(ByRef $Pattern) ; Also changes the default mode to utf
  Local Static $sMsg = "RegEx pattern error: "
  If Not StringRegExp($Pattern, "\(\*(?:ASC|UTF)\)") Then $Pattern = "(*UTF)" & $Pattern
  StringRegExp("", $Pattern, $STR_REGEXPARRAYMATCH) ; Test pattern on void input
  Switch @error
    Case 0
      If @extended = 1 Then
        cc($sMsg & $Pattern & @CRLF & Spaces(21) & "zero length match")
        Exit
      EndIf
    Case 2
      cc($sMsg & $Pattern & @CRLF & Spaces(20 + @extended) & "^")
      Exit
  EndSwitch
EndFunc   ;==>CheckPattern

;
; Core functions and use of $mArray ===================================

; #FUNCTION#
; Name ..........: _mRegEx
; Description ...: Internal helper function for mRegEx.
; Author ........: Bert Kerkhof

Func _mRegEx(Const $sInput, $Pattern, Const $iStart = 1)
  Local $aCap = StringRegExp($sInput, $Pattern, $STR_REGEXPARRAYFULLMATCH, $iStart)
  If @error Then Return Array("", $iStart, 0, $iStart)
  Local $iFound = @extended, $nCap = UBound($aCap) - 1
  Local Static $mArray[64] ; Optimized for execution speed
  If $nCap Then
    For $I = 1 To $nCap
      $mArray[$I] = $aCap[$I]
    Next
  Else
    $mArray[1] = $aCap[0]
    $nCap += 1
  EndIf
  $mArray[$nCap + 1] = $iStart ; $mSTART
  $mArray[$nCap + 2] = $iFound - StringLen($aCap[0]) ; $mFOUND
  $mArray[$nCap + 3] = $iFound ; $mOFFSET
  $mArray[0] = $nCap + 3
  Return $mArray
EndFunc   ;==>_mRegEx

; #FUNCTION#
; Name ..........: mRegEx
; Description ...: Search for a match, return captured data and properties
; Features:        1. This function searches for a match and returns
;                     captured data plus three properties.
;                  2. These properties reveal inside information about the
;                     completed search.
;                  3. The properties are named mSTART, mOFFSET and mFOUND.
;                  4. With a basic pattern, one data element is captured:
;                     the full match. In that case, the full match is
;                     returned in $mArray[1] and the three properties are
;                     in $mArray[2], $mArray[3] and $mArray[4].
;                  5. Advanced patterns may designate one or more capture
;                     fields, these are fragments within the match.
;                  6  The first element numbers in $mArray correspond
;                     exactly with the RegEx numbering of captures.
;                  7. The total number of captured fields is: $mArray[0]-3.
;                  8. Note: with one or more capture fields, the full
;                     match is not included in the results unless
;                     explicitly indicated as capture field in the
;                     pattern.
;                  9. If there is no match, $mArray only contains the
;                     three properties.
;                  10.All returned elements are always of the expected
;                     data-type. Finding the right RegEx pattern may be
;                     difficult, but the absence of exceptional behavior
;                     of returned elements will facilitate the journey.
; Syntax ........: mRegEx(Const $sInput, Const $Pattern[, $iStart = 1])
; Parameters ....: $sInput ..... String to search in.
;                  $Pattern .... RegEx search criteria.
;                  $iStart ..... [optional] location in $sInput where the
;                                search starts. Default is 1.
; Returns .......: Droste array. The first elements are the captured fields.
;                  The last three elements contain information about the
;                  executed search:
;                    mSTART .... Location where the search started.
;                    mFOUND .... Points to the first character of the match.
;                                A zero is returned if there is no match.
;                    mOFFSET ... Points to first character after the match.
;                                If there is no match, its value is mSTART.
;                  The total number of captured fields is $mArray[0]-3. If
;                  there is no match, the Droste array only contains the
;                  RegEx properties and the Droste COUNT element $mArray[0]
;                  is three.
; Author ........: Bert Kerkhof

Func mRegEx(Const $sInput, $Pattern, Const $iStart = 1)
  CheckPattern($Pattern)
  Return _mRegEx($sInput, $Pattern, $iStart)
EndFunc   ;==>mRegEx

; #FUNCTION#
; Name ..........: mRegExdotHeader
; Description ...: Retrieves the header of one search from $sInput
; Remark ........: Header is the text starting at mSTART up to mFOUND.
; Syntax ........: mRegExdotHeader(Const $mArray, Const $sInput)
; Parameters ....: $mArray ... [const] Map with captured data and properties
;                  $sInput ... [const] String that is searched.
; Returns .......: Header string
;                  If there is no match or there is no header, a null
;                  string ("") is returned.
; Example .......: aRegExSplit() and aRegExReplace()
; Author ........: Bert Kerkhof

Func mRegExdotHeader(Const $mArray, Const $sInput)
  Local Enum $mSTART = $mArray[0] - 2, $mFOUND
  Return StringMid($sInput, $mArray[$mSTART], $mArray[$mFOUND] - $mArray[$mSTART])
EndFunc   ;==>mRegExdotHeader

; #FUNCTION#
; Name ..........: mRegExdotFooter
; Description ...: Retrieves the footer from $sInput.
; Remakr ........: Footer is text starting at mOFFSET up to end of $sInput.
; Syntax ........: mRegExdotFooter(Const $mArray, Const $sInput)
; Parameters ....: $sInput ... [const] The string that is searched
;                  $mArray ... [const] Map with captured data and attributes
; Returns .......: Footer string
;                  If there is no match or there is no footer, a null
;                  string ("") is returned.
; Example .......: aRegExSplit() and aRegExReplace()
; Author ........: Bert Kerkhof

Func mRegExdotFooter(Const $mArray, Const $sInput)
  Return StringMid($sInput, $mArray[$mArray[0]])
EndFunc   ;==>mRegExdotFooter

; #FUNCTION#
; Name ..........: mRegExdotContent
; Description ...: Retrieves the contents of found captures as one string
; Syntax ........: mRegExdotContent(Const $mArray[, $sSep = "|"])
; Parameters ....: $mArray ... [const] Map with captured data.
;                  $Sep ...... [optional] If $Sep is a string, multiple
;                              capture fields (if any) will be separated
;                              by this string.
;                              If $Sep is a number, the capture fields are
;                              right justified with $Sep column width.
;                              Default is "|".
; Returns .......: Contents string
;                  If there is no match, a null string ("") is returned.
; Remark ........: If there are multiple capture fields, they are combined
;                  in one string, separated by the separation character.
; Author ........: Bert Kerkhof

Func mRegExdotContent(Const $mArray, $Sep = "|")
  Return sRecite(aLeft($mArray, $mArray[0] - 3), $Sep)
EndFunc   ;==>mRegExdotContent

; #FUNCTION#
; Name ..........: mRegExdotFound
; Description ...: Return position of the first character of the full match
; Syntax ........: mRegExdotFound(Const $mArray)
; Parameter .....: $mArray ... [const] Map with captured data and properties
; Returns .......: Location number pointing to a character in $sInput
;                  If there is no match, zero is returned
; Author ........: Bert Kerkhof

Func mRegExdotFound(Const $mArray)
  Return $mArray[$mArray[0] - 1]
EndFunc   ;==>mRegExdotFound

; #FUNCTION#
; Name ..........: mRegExdotLast
; Description ...: Returns position of the last character of the full match
; Syntax ........: mRegExdotLast(Const $mArray)
; Parameter .....: $mArray ... [const] Map with captured data and properties
; Returns .......: Location number pointing to a character in $sInput
;                  If there is no match, zero is returned
; Author ........: Bert Kerkhof

Func mRegExdotLast(Const $mArray)
  Local $mOFFSET = $mArray[0], $mFOUND = $mOFFSET - 1
  Return $mArray[$mFOUND] ? $mArray[$mOFFSET] - 1 : 0
EndFunc   ;==>mRegExdotLast

; #FUNCTION#
; Name ..........: mRegExdotLength
; Description ...: Returns the length of the full match
; Syntax ........: mRegExdotLength(Const $mArray)
; Parameter .....: $mArray ... [const] Map with captured data and properties
; Returns .......: Length number
;                  If there is no match, zero is returned
; Author ........: Bert Kerkhof

Func mRegExdotLength(Const $mArray)
  Local $mOFFSET = $mArray[0], $mFOUND = $mOFFSET - 1
  Return $mArray[$mFOUND] ? $mArray[$mOFFSET] - $mArray[$mFOUND] : 0
EndFunc   ;==>mRegExdotLength

;
; Functions that find a single match: =================================

; #FUNCTION#
; Name ..........: iRegEx
; Description ...: Search for a match. return the found location in $sInput
; Syntax ........: iRegEx(Const $sInput, Const $Pattern[, $iStart = 1])
; Parameters ....: $sInput.... String to search in.
;                  $Pattern... RegEx search pattern.
;                  $iStart.... [optional] location in $sInput where the
;                              search starts. Default is 1.
; Returns .......: If there is a match, the function returns the location of
;                  match found in the input string. This number can also be
;                  interpreted as logic (True). Otherwise returns 0 (False).
; AutoIT bonus ..: @extended points to last location of the match plus 1.
;                  If there is no match, @extended is zero.
; Author ........: Bert Kerkhof

Func iRegEx(Const $sInput, Const $Pattern, $iStart = 1)
  Local $mArray = mRegEx($sInput, $Pattern, $iStart)
  SetExtended($mArray[$mArray[0]]) ; mOFFSET
  Return $mArray[$mArray[0] - 1] ; mFOUND
EndFunc   ;==>iRegEx

; #FUNCTION#
; Name ..........: iRegExLast
; Description ...: Return the location of the last character of the match
; Syntax ........: iRegExLast(Const $sInput, Const $Pattern[, $iStart = 1])
; Parameters ....: $sInput.... String to search in.
;                  $Pattern... RegEx search pattern.
;                  $iStart.... [optional] location in $sInput where the
;                              search starts. Default is 1.
; Returns .......: If there is a match, the function returns the location of
;                  the last character of match found in the input string.
;                  If there is no match, zero is returned.
; Remark ........: This function is an alternative for usage of @extended
;                  in function iRegEx.
; Author ........: Bert Kerkhof

Func iRegExLast(Const $sInput, Const $Pattern, $iStart = 1)
  Local $mArray = mRegEx($sInput, $Pattern, $iStart)
  Return Max(0, $mArray[$mArray[0]] - 1)
EndFunc   ;==>iRegExLast

; #FUNCTION#
; Name ..........: sRegEx
; Description ...: Search for a match and return the full match
; Syntax ........: sRegEx(Const $sInput, $Pattern[, $iStart = 1)
; Parameters ....: $sInput ... String to search in.
;                  $Pattern .. RegEx search criteria.
;                  $iStart ... [optional] location in $sInput where the
;                              search starts. Default is 1.
; Returns .......: Full match string
; AutoIT bonus ..: Snapchat variable @extended points to first location of
;                  the match. If there is no match, @extended is zero.
; Author ........: Bert Kerkhof

Func sRegEx(Const $sInput, $Pattern, $iStart = 1)
  $Pattern = LimitCap($Pattern) ; Ignore capture fields, ensure full match
  Local $mArray = mRegEx($sInput, $Pattern, $iStart)
  Local $sOut = $mArray[1]
  SetExtended($mArray[$mArray[0] - 1])
  Return $sOut
EndFunc   ;==>sRegEx

;
; Functions to operate multiple search matches =====================

; #FUNCTION#
; Name ..........: nRegEx
; Description ...: Returns the total number of matches
; Syntax ........: nRegEx(Const $sInput, $Pattern[, $iStart = 1])
; Parameters ....: $sInput ... String to search in.
;                  $Pattern .. RegEx search pattern.
;                  $iStart ... [optional] location in $sInput where search
;                              starts. Default is 1.
; Returns .......: Total number of matches
; Author ........: Bert Kerkhof

Func nRegEx(Const $sInput, $Pattern, Const $iStart = 1)
  CheckPattern($Pattern)
  $Pattern = LimitCap($Pattern) ; Improves processing speed. No cons.
  Local $aCap = StringRegExp($sInput, $Pattern, $STR_REGEXPARRAYGLOBALMATCH, $iStart)
  Return @error ? 0 : UBound($aCap)
EndFunc   ;==>nRegEx

; #FUNCTION#
; Name ..........: amRegExCapture
; Description ...: Capture all data elements and properties of a multiple RegEx search
; Features ......: 1. Where mRegEx() finds one match, aRegExCapture finds
;                     all matches.
;                  2. The function always returns an array-in-array.
;                  3. The total number of matches is returned in the zero
;                     element of the returned array.
;                  4. If there is a match, element [1] contains a row of
;                     captured data and properties of the match.
;                  5. These properties reveal information about the location
;                     in the input where the match as found: The properties
;                     are named: mSTART, mOFFSET and  mFOUND.
;                     × Length of the full match can be calculated as
;                     $mFOUND minus $mOFFSET
;                     × Header (or spill) of the match is the input $mSTART
;                       up to $mFOUND
;                     × Footer of the last match is the input from $mOFFSET
;                       to end-of-file
;                  6. For next matches, additional rows are generated.
;                  7. There is no programmed maximum for the number of
;                     matches that the function can return.
;                  8. The Droste array-in-array technique allow for variable
;                     numbers of captured fields. For execution speed
;                     reasons, the maximum number of capture fields is set
;                     to 63. A higher maximum is easily programmable.
; Syntax ........: amRegExCapture(Const $sInput, $Pattern[, $iStart = 1])
; Parameters ....: $sInput ... String to search in.
;                  $Pattern .. RegEx search pattern.
;                  $iStart ... [optional] location in $sInput where search
;                              starts. Default is 1.
; Returns .......: Droste array-in-array with captured data and properties.
;                  If there is no match, the COUNT property of the main
;                  array is zero.
; Author ........: Bert Kerkhof

Func amRegExCapture(Const $sInput, $Pattern, $iStart = 1)
  CheckPattern($Pattern)
  ; Optimized for execution speed:
  Local $mArray, $iOffset, $amResult = aNew(), $iStrek = 1
  While True
    For $I = $iStrek To UBound($amResult) - 1
      $iOffset = $iStart
      $mArray = _mRegEx($sInput, $Pattern, $iOffset)
      $iStart = $mArray[$mArray[0]]
      If $iStart = $iOffset Then ExitLoop
      $amResult[$I] = $mArray
    Next
    If $iStart = $iOffset Then ExitLoop
    $iStrek = Ubound($amResult)
    _NewDim($amResult, $iStrek)
  WEnd
  $amResult[0] = $I - 1
  Return $amResult
EndFunc   ;==>amRegExCapture

; #FUNCTION#
; Name ..........: aRegExLocate
; Description ...: Returns a one-dimensional array with multiple findings
; Features ......: For each match, the location element $mFOUND-1 is returned.
;                  This points to the last character in the match.
;                  Optionally, the $mOFFSET-1 location can be returned.
;                  The total number of matches is in the zero element.
; Syntax ........: aRegExLocate(Const $sInput, $Pattern[, $iStart = 1[, $Lfirst = True]])
; Parameters ....: $sInput ... String to search in.
;                  $Pattern .. RegEx search pattern.
;                  $iStart. .. [optional] location in $sInput where search
;                              starts.
; Returns .......: Droste array of numbers which point to the locations of
;                  the full matches.
; Author ........: Bert Kerkhof

Func aRegExLocate(Const $sInput, $Pattern, Const $iStart = 1, Const $Lfirst = True)
  CheckPattern($Pattern)
  $Pattern = LimitCap($Pattern) ; Improves processing speed. No cons.
  Local $amSource = amRegExCapture($sInput, $Pattern, $iStart)
  Local $aResult = aNew($amSource[0])
  Local Enum $mFOUND = 3, $mOFFSET
  If $Lfirst Then
    For $I = 1 To $amSource[0]
      $aResult[$I] = ($amSource[$I])[$mFOUND]
    Next
  Else
    For $I = 1 To $amSource[0]
      $aResult[$I] = ($amSource[$I])[$mOFFSET] - 1
    Next
  EndIf
  Return $aResult
EndFunc   ;==>aRegExLocate

; #FUNCTION#
; Name ..........: aRegExSplit
; Description ...: Splits the sInput string in parts
;                  Resembles the StringSplit() function but offers more
;                  choice in separators.
; Features ......: Another difference with StringSplit is that an empty
;                  sInput string result in zero parts.
;                  Even if zero parts are returned, the returned value
;                  is of type array.
; Syntax ........: aRegExSplit($sInput, $Pattern[, $iStart = 1])
; Parameters: ...: $sInput ... String to search in.
;                  $Pattern .. RegEx search criteria.
;                  $iStart ... [optional] up to this location, the $sInput
;                              is ignored. Default is no ignoring.
; Returns: ......: Array with the parts. The count property of the array
;                  contains the number of parts.
; Author ........: Bert Kerkhof

Func aRegExSplit(Const $sInput, $Pattern, Const $iStart = 1)
  Local Static $mOFFSET = 4
  Local $amSource = amRegExCapture($sInput, LimitCap($Pattern), $iStart)
  If $amSource[0] = 0 Then Return Array($sInput) ; Unsplitted
  Local $mArray, $aSplit = aNew($amSource[0])
  For $I = 1 To $amSource[0]
    $mArray = $amSource[$I]
    $aSplit[$I] = mRegExdotHeader($mArray, $sInput)
  Next
  If $mArray[$mOFFSET] <= StringLen($sInput) Then aAdd($aSplit, mRegExdotFooter($mArray, $sInput))
  Return $aSplit
EndFunc   ;==>aRegExSplit

; #FUNCTION#
; Name ..........: sRegExReplace
; Description ...: Replaces all matches with the outcome of a script
; Features ......: 1. Finds all matches and replaces these by the outcome of
;                     a script.
;                  2. With capture fields in the pattern and placeholders in
;                     the script,
;                     parts of the match may return in the replacements.
;                  3. If for a back-reference in the script is no
;                     corresponding capture field in the pattern, an
;                     informative message is issued at the console and the
;                     program halts.
;                  4. If needed, the function can convert a backreference to
;                     uppercase.
; Syntax ........: sRegExReplace($sInput, $Pattern, $sScript[, $iStart = 1])
; Parameters ....: $sInput .... String to search in
;                  $Pattern ... RegEx pattern
;                  $sScript ... Describes the replacement. May include
;                               back-references and call-outs.
;                  $iStart .... [optional] location in $sInput where search
;                               starts.
; Returns .......: String with all the replacements made.
; AutoIT bonus ..: Snapchat variable @extended contains the total number
;                  of matches.
; Explanation ...: Use of this function needs some understanding of the
;                  'capture fields' mechanism. The pattern may contain one
;                  or more sub matches called capture fields. These are
;                  parts of the full match. The function replaces full
;                  matches with outcome of the script, not just replace a
;                  captured field. The contents of the captures - if any –
;                  can be inserted anywhere in the script by means of
;                  placeholders, the so called 'back-references'.
; Back-references:
;     Placeholders in the script are replaced with what is captured:
;       \1 references the first captured field. \2 the second, etcetera.
;       Placeholder $ gives the same result.
;       Placeholder ^ causes a conversion of the capture to uppercase.
; Author ........: Bert Kerkhof

Func sRegExReplace(Const $sInput, $Pattern, Const $sScript, Const $iStart = 1)

  ; Prepare the replace tokens:
  Local $aaToken = aRegExCapture($sScript, "(?<!\\)(((?:\\|\$|\^))(\d+))")
  Local Enum $pHOLDER = 1, $pSYMBOL, $pNUMBER ; $aToken elements
  Local $aToken, $ePos, $nCap = nCapture($Pattern) ; Number of capture fields
  For $I = 1 To $aaToken[0] ; Check validity of the numbers:
    $aToken = $aaToken[$I]
    If $aToken[$pNUMBER] <= $nCap Then ContinueLoop
    $ePos = StringInStr($sScript, $aToken[1])
    cc('sReplace pattern error: ' & $sScript & @CRLF & Spaces($ePos + 24) & "^")
    Exit
  Next

  ; Capture headers, fragments and one footer:
  If $aaToken[0] = 0 Then $Pattern = LimitCap($Pattern)
  Local $amSource = amRegExCapture($sInput, $Pattern, $iStart)
  If $amSource[0] = 0 Then Return SetError(0, 0, $sInput)
  Local $mArray, $sBpart, $sTarget, $sOut = ""
  For $I = 1 To $amSource[0]
    $mArray = $amSource[$I]
    $sOut &= mRegExdotHeader($mArray, $sInput)
    $sBpart = $sScript
    For $J = 1 To $aaToken[0]
      $aToken = $aaToken[$J]
      $sTarget = $mArray[$aToken[$pNUMBER]]
      If $aToken[$pSYMBOL] = "^" Then $sTarget = StringUpper($sTarget)
      $sBpart = StringReplace($sBpart, $aToken[$pHOLDER], $sTarget)
    Next
    $sOut &= $sBpart
  Next
  $sOut &= mRegExdotFooter($mArray, $sInput)
  SetExtended($amSource[0])
  Return $sOut
EndFunc   ;==>sRegExReplace

; #FUNCTION#
; Name ..........: aRegExCapture
; Description ...: Capture all data elements of a multiple RegEx search
; Remark ........: In some instances, the location properties that
;                  amRegExCapture reveal, are not needed. In that case a
;                  more simple implementation might show execution speed
;                  advantage. The aRegExCapture function is coded to
;                  compare. Note that the returned array can have one or
;                  two dimensions, depending on input pattern. In general
;                  it is good practice to avoid a function that can returns
;                  confusing different data structures.
; Features ......: 1. If the pattern has zero capture fields, each returned
;                     Droste array element contains the full match as
;                     string.
;                  2. If the pattern has one capture field, each element
;                     contains the capture as string.
;                  3. If the pattern has more capture fields, the function
;                     returns a Droste array-in-array. Each row starts with
;                     the zero element which contains the count of captured
;                     fields.
;                  4  The element numbers in the row correspond exactly with
;                     the RegEx numbering of captures.
;                  5. Note: with one or more capture fields, a full match is
;                     not included in the results unless explicitly
;                     indicated as capture field in the pattern.
; Syntax ........: aRegExCapture(Const $sInput, $Pattern[, $iStart = 1])
; Parameters ....: $sInput ... String to search in.
;                  $Pattern .. RegEx search pattern.
;                  $iStart ... [optional] location in $sInput where search
;                              starts. Default is 1.
; Returns .......: Droste array-in-array with captured data and properties.
; Author ........: Bert Kerkhof

Func aRegExCapture(Const $sInput, $Pattern, Const $iStart = 1)
  CheckPattern($Pattern)
  Local $aaMatrix, $aCap, $nCap = nCapture($Pattern) ; Number of captures
  $aCap = StringRegExp($sInput, $Pattern, $STR_REGEXPARRAYGLOBALMATCH, $iStart)
  If @error Then Return aNew()
  If $nCap = 1 Then
    $aaMatrix = aDroste($aCap)
  Else
    Local $N = UBound($aCap) / $nCap
    $aaMatrix = aNew($N)
    Local $aColumn = aNew($nCap), $iStep = -1
    For $iRow = 1 To $N
      For $iColumn = 1 To $nCap
        $aColumn[$iColumn] = $aCap[$iStep + $iColumn]
      Next
      $aaMatrix[$iRow] = $aColumn
      $iStep += $nCap
    Next
  EndIf
  Return $aaMatrix
EndFunc   ;==>aRegExCapture

; #FUNCTION#
; Name ..........: sTitleCase
; Description ...: Modifies the first character of the words in the string
;                  Alternative to StringTitleCase
; Syntax ........: StringTitleCase($sTitle)
; Parameters ....: $sTitle ..... String to modify
; Returns .......: Modified string
; Author ........: Bert Kerkhof

Func sTitleCase(Const $sInput)
  Return sRegExReplace($sInput, "(^|\h+)(\p{Ll})", "$1^2")
EndFunc   ;==>sTitleCase

; End ===============================================================
