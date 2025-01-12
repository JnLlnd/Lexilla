CvtClr(Color) {
    Return (Color & 0xFF) << 16 | (Color & 0xFF00) | (Color >>16)
}

GetErrorMessage(ErrorCode, LanguageId := 0) {
    VarSetCapacity(ErrorMsg, 8192)
    DllCall("Kernel32.dll\FormatMessage"
        , "UInt", 0x1200 ; FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS
        , "UInt", 0
        , "UInt", ErrorCode + 0
        , "UInt", LanguageId
        , "Str" , ErrorMsg
        , "UInt", 8192)
    Return StrGet(&ErrorMsg)
}
