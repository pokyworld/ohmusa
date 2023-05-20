<!--#include virtual="/shop/payment/functions/helpers.inc"-->

<%
input = "This is the BBC in london calling"
RW(input)
aryWords = Split(input," ")
For w = 0 To UBound(aryWords)
  newWord = ""
  newWord = UCase(Left(aryWords(w), 1)) & Right(aryWords(w), Len(aryWords(w))-1)
  RW(aryWords(w) & " : " & newWord)
  aryWords(w) = newWord
Next
output = Join(aryWords, " ")
RW(output)
%>