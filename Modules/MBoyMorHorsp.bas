Attribute VB_Name = "MBoyMorHorsp"
Option Explicit

Sub Debug_Print(ByVal s As String)
    Form1.Text1 = Form1.Text1 & s & vbCrLf
End Sub

Function find(haystack() As Byte, needle() As Byte, Optional start As Long = 0) As Long
    'Boyer-Moore-Horspool-Algo
    find = -1
    Dim i_n As Long
    Dim i_h As Long:    i_h = start
    Const UCHAR_MAX As Long = 255
    Dim bad_char_skip(0 To UCHAR_MAX + 2) As Long
    On Error GoTo 0
    Dim ds As String
    Dim n_h As Long: n_h = UBound(haystack) + 1
    Dim n_n As Long: n_n = UBound(needle) + 1
    If n_n > n_h Then Exit Function
    
    For i_n = 0 To UCHAR_MAX + 2
        bad_char_skip(i_n) = n_n
    Next
    Dim last As Long: last = n_n - 1
    For i_n = 0 To last - 1
        bad_char_skip(needle(i_n)) = last - i_n
    Next
    Dim bcs As Byte, bhs As Byte
    'Dim ds As String 'debugstring
    Debug_Print "We search the haystack from the left, but "
    Debug_Print "we compare with each character of the needle from the right"
    While (n_h - start) >= n_n
        i_n = last
        While haystack(i_h + i_n) = needle(i_n)
            i_n = i_n - 1
            If i_n = 0 Then
                find = i_h
                Exit Function
            End If
        Wend
        bhs = haystack(i_h + last)
        ds = """" & Chr(bhs) & """"
        bcs = bad_char_skip(bhs)
        If bcs < (last + 1) Then
            ds = ds & ": we are allowed to skip minimum " & bcs & " characters "
        Else
            ds = ds & " is not in the needle so we skip the whole length of the needle: " & bcs
        End If
        Debug_Print ds
        n_h = n_h - bcs 'bad_char_skip(haystack(i_h + last))
        i_h = i_h + bcs 'bad_char_skip(haystack(i_h + last))
        Debug_Print "n_h: " & n_h & "    i_h: " & i_h & "    bcs: " & bcs
    Wend
End Function

'nop
Function findX(haystack() As Byte, needle() As Byte, Optional start As Long = 0) As Long
    findX = -1
    Dim u_h As Long: u_h = UBound(haystack) '+ 1
    Dim u_n As Long: u_n = UBound(needle) '+ 1
    Dim i_n As Long
    Dim i_h As Long: i_h = start
    While i_h < (u_h - u_n)
        i_n = u_n
        While haystack(i_h + i_n) = needle(i_n)
            i_n = i_n - 1
            If i_n = 0 Then
                findX = i_h
                Exit Function
            End If
        Wend
        i_h = i_h + u_n + 1
    Wend
End Function
'Der Algo läuft ganz gut wenn man eine Nadel und einen Heuhaufen als eine Bytefolge vorliegen hat.
'Der Algo legt ein Feld an, der Größe UCHAR_MAX.
'Die meisten Beispiele die man zu dem Algo findet hantieren mit Strings weil es sehr anschaulich wirkt.
'Zu Zeiten von 16-bit-Windows war das OK, Strings waren damals noch wirklich Bytefolgen.
'Heute ist Unicode, und wir haben gelernt daß die Größe eines Characters nicht festgelegt ist.
'wie soll man ein skip-Array anlegen wenn man die Größe eines Zeichens nicht kennt?
'Ist das nicht eine Katastrophe für den Algo?
'man muß ein Feld anlegen das einen Wert für jedes Byte speichert der Größe UCHAR_MAX
'man kann jetzt damit argumentieren, daß jeder String auch als Stream von Bytes betrachtet werden kann.
'OK, aber je nach Encoding hat der Algo eine sehr viel schlechtere Laufzeit.
'wenn bspw mit UTF16 oder UCS nahezu jedes zweite Byte eine 0 ist, dann dürfen immer nur maximal 2 Bytes geskippt werden.
'
'

'code aus wikipedia
'boyermoore_horspool_memmem(const unsigned char* haystack, ssize_t hlen,
'                           const unsigned char* needle,   ssize_t nlen)
'{
'    size_t scan = 0;
'    size_t bad_char_skip[UCHAR_MAX + 1]; /* Officially called:
'                                          * bad character shift */
'
'    /* Sanity checks on the parameters */
'    if (nlen <= 0 || !haystack || !needle)
'        return NULL;
'
'    /* ---- Preprocess ---- */
'    /* Initialize the table to default value */
'    /* When a character is encountered that does not occur
'     * in the needle, we can safely skip ahead for the whole
'     * length of the needle.
'     */
'    for (scan = 0; scan <= UCHAR_MAX; scan = scan + 1)
'        bad_char_skip[scan] = nlen;
'
'    /* C arrays have the first byte at [0], therefore:
'     * [nlen - 1] is the last byte of the array. */
'    size_t last = nlen - 1;
'
'    /* Then populate it with the analysis of the needle */
'    for (scan = 0; scan < last; scan = scan + 1)
'        bad_char_skip[needle[scan]] = last - scan;
'
'    /* ---- Do the matching ---- */
'
'    /* Search the haystack, while the needle can still be within it. */
'    While (hlen >= nlen)
'    {
'        /* scan from the end of the needle */
'        for (scan = last; haystack[scan] == needle[scan]; scan = scan - 1)
'            if (scan == 0) /* If the first byte matches, we've found it. */
'                return haystack;
'
'        /* otherwise, we need to skip some bytes and start again.
'           Note that here we are getting the skip value based on the last byte
'           of needle, no matter where we didn't match. So if needle is: "abcd"
'           then we are skipping based on 'd' and that value will be 4, and
'           for "abcdd" we again skip on 'd' but the value will be only 1.
'           The alternative of pretending that the mismatched character was
'           the last character is slower in the normal case (Eg. finding
'           "abcd" in "...azcd..." gives 4 by using 'd' but only
'           4-2==2 using 'z'. */
'        hlen     -= bad_char_skip[haystack[last]];
'        haystack += bad_char_skip[haystack[last]];
'    }
'
'    return NULL;
'}
'
'

