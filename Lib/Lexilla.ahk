LoadLexilla(DllPath := "Lexilla.dll", Notify := "__SciNotify") {
    If (!DllCall("LoadLibrary", "Str", DllPath, "Ptr")) {
        Return 0
    }

    Return 1
}

