class ContextException : Exception {
    [string] $AdditionalData

    ContextException($Message, $AdditionalData) : base($Message) {
        $this.AdditionalData = $AdditionalData
    }
}
