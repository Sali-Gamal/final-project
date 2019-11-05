function AddRecord() {
    var adoConn = new ActiveXObject("ADODB.Connection");
    var adoRS = new ActiveXObject("ADODB.Recordset");
    
    adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='/\DoctorClinic.accdb'");
    adoRS.Open("Select * From PatientTable", adoConn, 1, 3);
    
    adoRS.AddNew;
    adoRS.Fields("PatientID").value = 5;
    adoRS.Update;
    
    adoRS.Close();
    adoConn.Close();
    }  

    function AddRecord2() {
        var cn = new ActiveXObject("ADODB.Connection");
               var strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\\Users\\shadi\\web\\final-project\\DoctorClinic.accdb";
               cn.Open(strConn);
               var rs = new ActiveXObject("ADODB.Recordset");
               var SQL = "select * from PatientTable";
               rs.Open(SQL, cn);
               alert(rs(0));
               rs.AddNew
               rs.Fields("PatientID") = 7;
               rs.Update;   
               rs.Close();
               cn.Close(); 
       
       }