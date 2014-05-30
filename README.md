OutputExcel
===========

output  List&lt;T> To excel then download 

Asp.net List<T> conversation of Excel then give client to download



EXAMPLE  
====================================================================
Here is asp.net button
====================================================================

    private void btnExcelReport_Click(object sender, EventArgs e)
    {
        
        var data = ProductTestingDao.GetAll();
        
        OutputExcel.ResponseExcel(this.Response, data);
    }
