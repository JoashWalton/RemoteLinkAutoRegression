#include <Array.au3>

func genPass($max_len)

   Local $aPass[$max_len]

   For $i = 0 To UBound($aPass) - 1
	   $aPass[$i] = chr(Random(33, 126, 1))
   Next

   $aPass = StringReplace(_ArrayToString($aPass, @TAB, 1, $max_len), @TAB, "")

   Return $aPass

EndFunc

; Generate a pseudo-random password of 8 characters in length, consisting of alpha
; numeric and symbolic characters, including mixed-case letters.

; For DEMO purposes:
; Generate a strong, 12 char long password

   msgbox(0,"Your new password", "Your new pseudo-random password is: " & genPass(32))