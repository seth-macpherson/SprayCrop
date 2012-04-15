<%@ Page Language="vb" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<%  Server.ScriptTimeout = 1000%>

<html>

<Script Language="VB" RunAt="Server">

    Sub Page_Load(ByVal Sender As Object, ByVal e As EventArgs)
  
    End Sub

    Sub Upload_Click(ByVal Sender As Object, ByVal e As EventArgs)

        ' Display properties of the uploaded file

        FileName.InnerHtml = MyFile.PostedFile.FileName
        FileContent.InnerHtml = MyFile.PostedFile.ContentType
        FileSize.InnerHtml = MyFile.PostedFile.ContentLength
        UploadDetails.Visible = True

        ' Let us recover only the file name from its fully qualified path at client 

        Dim strFileName As String
        
        strFileName = MyFile.PostedFile.FileName
        
        Dim c As String = System.IO.Path.GetFileName(strFileName) ' only the attched file  name not its path

        If MyFile.PostedFile.ContentLength <= 100000 And _
            (MyFile.PostedFile.ContentType = "image/gif" Or _
                MyFile.PostedFile.ContentType = "image/jpeg" Or _
                MyFile.PostedFile.ContentType = "image/pjpeg") Then

            Try

                MyFile.PostedFile.SaveAs(Request.PhysicalApplicationPath + "\logos\p" + Request.QueryString("pnum") + Right(c, 4))

                Dim cn As SqlConnection
                Dim cm As SqlCommand
            
                cn = New SqlConnection("SERVER=GTSERVER;UID=unison;PWD=agspray08;DATABASE=agspray;")
                cn.Open()
                cm = New SqlCommand("update packers set logofileext='" + Right(c, 4) + "' where packernumber='" + Request.QueryString("pnum") + "'", cn)
                cm.ExecuteNonQuery()
                cm.Dispose()
                cn.Close()
                
                Span1.InnerHtml = "<" + "script" + ">window.opener.document.forms[1].update.click();self.close();<" + "/script" + ">"
                
                '"Your logo was successfully uploaded.<br><br><input type=button name=btnclose onclick=""window.opener.document.forms[1].update.click();self.close();"" value=""Close Window"" />"

            Catch Exp As Exception
         
                Span1.InnerHtml = "An Error occured. Please check the attached file and try again."
                Span1.InnerHtml = Exp.ToString
                UploadDetails.Visible = False
                Span2.Visible = False
            
            End Try
        
        Else
            
            Span1.InnerHtml = "Logo file does not meet requirements. Please check the attached file and try again."
            
        End If
        
    End Sub
       
</Script>
   
<Body>
      
         <Font Color="Navy" Face=Helvetica Size=5> <B>Logo Upload
         </Font>
         
                  
         <br />-
          <Font Color="Navy" Face=Helvetica Size=1>GIF, JPEG
          </font>
         <br />-
          <Font Color="Navy" Face=Helvetica Size=1>No Larger than 100kb, and within 250x50 dimensions
          </font>
         
          <HR Size="2" Color=Black >
          <P>
                   
         <Form Method="Post" EncType="Multipart/Form-Data" RunAt="Server">

         Choose Your File To Upload : <BR>
         <Input ID="MyFile" Type="File" RunAt="Server" Size="40"> <BR>
         <BR>

         <Input Type="Submit" Value="Upload" OnServerclick="Upload_Click" RunAt="Server">

         <P>
         <Div ID="UploadDetails" Visible="False" RunAt="Server">
         
            File Name: <Span ID="FileName" RunAt="Server"/> <BR>
            File Content: <Span ID="FileContent" RunAt="Server"/> 

           <BR>
         
            File Size: <Span ID="FileSize" RunAt="Server"/>bytes
         
           <BR></Div>

           <Span ID="Span1" Style="Color:Red" RunAt="Server"/>
           <Span ID="Span2" Style="Color:Red" RunAt="Server"/>
         
      </Form>

</Body>

</html>

