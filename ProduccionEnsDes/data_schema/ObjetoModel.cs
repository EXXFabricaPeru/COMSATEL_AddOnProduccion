namespace AddonProduccionEnsDes.data_schema
{
    public class ObjetoModel
    {
        public ObjetoModel()
        {
            findColumns = null;
            childTables = null;
            formColumns = null;
            childFormColumns = null;
        }

        public string code { get; set; }
        public string name { get; set; }
        public string tableName { get; set; }
        public string[] findColumns { get; set; }
        public string[] childTables { get; set; }
        public SAPbobsCOM.BoYesNoEnum canCancel { get; set; }
        public SAPbobsCOM.BoYesNoEnum canClose { get; set; }
        public SAPbobsCOM.BoYesNoEnum canDelete { get; set; }
        public SAPbobsCOM.BoYesNoEnum canCreateDefaultForm { get; set; }
        public string[] formColumns { get; set; }
        public SAPbobsCOM.BoYesNoEnum canFind { get; set; }
        public SAPbobsCOM.BoYesNoEnum canLog { get; set; }
        public SAPbobsCOM.BoUDOObjType objectType { get; set; }
        public SAPbobsCOM.BoYesNoEnum manageSeries { get; set; }
        public SAPbobsCOM.BoYesNoEnum enableEnhancedForm { get; set; }
        public SAPbobsCOM.BoYesNoEnum rebuildEnhancedForm { get; set; }
        public string[] childFormColumns { get; set; }
    }
}
