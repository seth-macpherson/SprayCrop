	<script language="javascript">
		function calculateTotal(){
			if (document.frmadd.RateAcre.value == "")
				document.frmadd.RateAcre.value = 0
			if (document.frmadd.AcresTreated.value == "")
				document.frmadd.AcresTreated.value = 0
			if (document.frmadd.TotalMaterialApplied.value == "")
				document.frmadd.TotalMaterialApplied.value = 0
				
			if (!(isNaN(document.frmadd.AcresTreated.value))) {
				if (!(isNaN(document.frmadd.RateAcre.value)) && (document.frmadd.RateAcre.value != 0)) {
					document.frmadd.TotalMaterialApplied.value = (Math.round(100*document.frmadd.AcresTreated.value *  document.frmadd.RateAcre.value)/100) ;
				}
				else{
					if (!(isNaN(document.frmadd.TotalMaterialApplied.value)) && document.frmadd.TotalMaterialApplied.value != 0) {
						document.frmadd.RateAcre.value = (Math.round(100*document.frmadd.TotalMaterialApplied.value / document.frmadd.AcresTreated.value)/100) ;
					}
					else
					{
						//alert("Please enter a numeric value.");
						document.frmadd.RateAcre.value = 0;
						document.frmadd.RateAcre.focus();
					}
				}
			}
			else{
				//alert("Please enter a numeric value.");
				document.frmadd.AcresTreated.value = 0;
				document.frmadd.AcresTreated.focus();
			}
		}
		function prefilAcres() {
			if (document.frmadd.AcresTreated1)
			{
				document.frmadd.AcresTreated1.value = document.frmadd.AcresTreated.value;
				calculateTotal1();
			}
			if (document.frmadd.AcresTreated2)
			{
				document.frmadd.AcresTreated2.value = document.frmadd.AcresTreated.value;
				calculateTotal2();
			}
			if (document.frmadd.AcresTreated3)
			{
				document.frmadd.AcresTreated3.value = document.frmadd.AcresTreated.value;
				calculateTotal3();
			}
			if (document.frmadd.AcresTreated4)
			{
				document.frmadd.AcresTreated4.value = document.frmadd.AcresTreated.value;
				calculateTotal4();
			}
			if (document.frmadd.AcresTreated5)
			{
				document.frmadd.AcresTreated5.value = document.frmadd.AcresTreated.value;
				calculateTotal5();
			}
			document.frmadd.RateAcre.focus();
		}	
		function displaySprayListData(){
			for (i = 0; i < document.frmadd.SprayListID.options.length; i++) {
			   if (document.frmadd.SprayListID.options [i].selected == true) {
    				maxappvalues = document.frmadd.SprayListID.options [i].value.split("|");
					document.frmadd.MaxAppUse.value = maxappvalues[1];
					document.frmadd.MaxAppSeason.value = maxappvalues[2];
					document.frmadd.Units.value = maxappvalues[3];
  				 }
			}
		}
					
		function refreshdata(prodID){
			document.frmadd.RefreshData.value=1;
			document.frmadd.NewProdID.value = prodID;
			document.frmadd.submit();
		}

		function populateGrowerList(){
			//code from the EnterSprayData.asp
			//not in use, but saved here.
			//javascript code in SprayRecords_List identical except for this f(x).
			//that f(x) is below.
			window.alert('populateGrowerList in enterspraydata_js.txt');
			growerID = 0;
			for (i = 0; i < document.frmadd.LocationOptions.options.length; i++) {
				document.frmadd.LocationOptions.options [i] = null;
			}
			for (i = 0; i < document.frmadd.GrowerID.options.length; i++) {
			   if (document.frmadd.GrowerID.options [i].selected == true) {
 			 	    document.frmadd.LocationOptions.options [0] = new Option ("NEW LOCATION","", false, false);
    				locations = document.frmadd.GrowerID.options [i].value.split("|");
					growerID = locations[0];
					if (locations.length > 1){
						for (var j = 1; j < locations.length; j++) {
						   document.frmadd.LocationOptions.options [j] = new Option (
						   locations [j],
						   locations [j], false, false);
						  }
					}
  				 }
			}
<%
	set rsSelect = GetActiveGrowers()
	IF not rsSelect.EOF THEN
%>
			switch(parseInt(growerID)){
<%
		DO WHILE not rsSelect.eof 
%>
				case(<%=rsSelect.Fields("GrowerID")%>):
					document.frmadd.Supervisor.value = "<%=rsSelect.Fields("ApplicatorSupervisor")%>";
					document.frmadd.LicenseNumber.value = "<%=rsSelect.Fields("SupervisorLicense")%>";
					document.frmadd.Applicator.value = "<%=rsSelect.Fields("Applicator")%>";
					document.frmadd.ApplicatorLicense.value = "<%=rsSelect.Fields("ApplicatorLicense")%>";
					document.frmadd.ChemicalSupplier.value = "<%=rsSelect.Fields("ChemicalSupplier")%>";
					document.frmadd.RecommendedBy.value = "<%=rsSelect.Fields("RecommendedBy")%>";
					break;
<%
rsSelect.MoveNext
LOOP
%>
			}

<%
END IF	%>
		}
		
		function populateGrowerList_EDIT(){
			//see comments in above populateGrowerList().
			window.alert('populateGrowerList in sprayrecords_list_js.txt');
			for (i = 0; i < document.frmadd.LocationOptions.options.length; i++) {
				document.frmadd.LocationOptions.options [i] = null;
			}
			for (i = 0; i < document.frmadd.GrowerID.options.length; i++) {
			   if (document.frmadd.GrowerID.options [i].selected == true) {
 			 	    document.frmadd.LocationOptions.options [0] = new Option ("NEW LOCATION","", false, false);
    				locations = document.frmadd.GrowerID.options [i].value.split("|");
					if (locations.length > 1){
						for (var j = 1; j < locations.length; j++) {
						   document.frmadd.LocationOptions.options [j] = new Option (
						   locations [j],
						   locations [j], false, false);
						  }
					}
  				 }
			}
		}

		
		function calculateTotal5(){
			if (document.frmadd.RateAcre5.value == "")
				document.frmadd.RateAcre5.value = 0
			if (document.frmadd.AcresTreated5.value == "")
				document.frmadd.AcresTreated5.value = 0
			if (document.frmadd.TotalMaterialApplied5.value == "")
				document.frmadd.TotalMaterialApplied5.value = 0
				
			if (!(isNaN(document.frmadd.AcresTreated5.value))) {
				if (!(isNaN(document.frmadd.RateAcre5.value)) && (document.frmadd.RateAcre5.value != 0)) {
					document.frmadd.TotalMaterialApplied5.value = (Math.round(100*document.frmadd.AcresTreated5.value *  document.frmadd.RateAcre5.value)/100) ;
				}
				else{
					if (!(isNaN(document.frmadd.TotalMaterialApplied5.value)) && document.frmadd.TotalMaterialApplied5.value != 0) {
						document.frmadd.RateAcre5.value = (Math.round(100*document.frmadd.TotalMaterialApplied5.value / document.frmadd.AcresTreated5.value)/100) ;
					}
					else
					{
						//alert("Please enter a numeric value.");
						document.frmadd.RateAcre5.value = 0;
						document.frmadd.RateAcre5.focus();
					}
				}
			}
			else{
				//alert("Please enter a numeric value.");
				document.frmadd.AcresTreated5.value = 0;
				document.frmadd.AcresTreated5.focus();
			}
		}	
		function displaySprayListData5(){
			for (i = 0; i < document.frmadd.SprayListID5.options.length; i++) {
			   if (document.frmadd.SprayListID5.options [i].selected == true) {
    				maxappvalues = document.frmadd.SprayListID5.options [i].value.split("|");
					document.frmadd.MaxAppUse5.value = maxappvalues[1];
					document.frmadd.MaxAppSeason5.value = maxappvalues[2];
					document.frmadd.Units5.value = maxappvalues[3];
  				 }
			}
		}										


		function calculateTotal4(){
			if (document.frmadd.RateAcre4.value == "")
				document.frmadd.RateAcre4.value = 0
			if (document.frmadd.AcresTreated4.value == "")
				document.frmadd.AcresTreated4.value = 0
			if (document.frmadd.TotalMaterialApplied4.value == "")
				document.frmadd.TotalMaterialApplied4.value = 0
				
			if (!(isNaN(document.frmadd.AcresTreated4.value))) {
				if (!(isNaN(document.frmadd.RateAcre4.value)) && (document.frmadd.RateAcre4.value != 0)) {
					document.frmadd.TotalMaterialApplied4.value = (Math.round(100*document.frmadd.AcresTreated4.value *  document.frmadd.RateAcre4.value)/100) ;
				}
				else{
					if (!(isNaN(document.frmadd.TotalMaterialApplied4.value)) && document.frmadd.TotalMaterialApplied4.value != 0) {
						document.frmadd.RateAcre4.value = (Math.round(100*document.frmadd.TotalMaterialApplied4.value / document.frmadd.AcresTreated4.value)/100) ;
					}
					else
					{
						//alert("Please enter a numeric value.");
						document.frmadd.RateAcre4.value = 0;
						document.frmadd.RateAcre4.focus();
					}
				}

			}
			else{
				//alert("Please enter a numeric value.");
				document.frmadd.AcresTreated4.value = 0;
				document.frmadd.AcresTreated4.focus();
			}
		}	
		function displaySprayListData4(){
			for (i = 0; i < document.frmadd.SprayListID4.options.length; i++) {
			   if (document.frmadd.SprayListID4.options [i].selected == true) {
    				maxappvalues = document.frmadd.SprayListID4.options [i].value.split("|");
					document.frmadd.MaxAppUse4.value = maxappvalues[1];
					document.frmadd.MaxAppSeason4.value = maxappvalues[2];
					document.frmadd.Units4.value = maxappvalues[3];
  				 }
			}
		}										

		function calculateTotal3(){
			if (document.frmadd.RateAcre3.value == "")
				document.frmadd.RateAcre3.value = 0
			if (document.frmadd.AcresTreated3.value == "")
				document.frmadd.AcresTreated3.value = 0
			if (document.frmadd.TotalMaterialApplied3.value == "")
				document.frmadd.TotalMaterialApplied3.value = 0
				
			if (!(isNaN(document.frmadd.AcresTreated3.value))) {
				if (!(isNaN(document.frmadd.RateAcre3.value)) && (document.frmadd.RateAcre3.value != 0)) {
					document.frmadd.TotalMaterialApplied3.value = (Math.round(100*document.frmadd.AcresTreated3.value *  document.frmadd.RateAcre3.value)/100) ;
				}
				else{
					if (!(isNaN(document.frmadd.TotalMaterialApplied3.value)) && document.frmadd.TotalMaterialApplied3.value != 0) {
						document.frmadd.RateAcre3.value = (Math.round(100*document.frmadd.TotalMaterialApplied3.value / document.frmadd.AcresTreated3.value)/100) ;
					}
					else
					{
						//alert("Please enter a numeric value.");
						document.frmadd.RateAcre3.value = 0;
						document.frmadd.RateAcre3.focus();
					}
				}
			}
			else{
				//alert("Please enter a numeric value.");
				document.frmadd.AcresTreated3.value = 0;
				document.frmadd.AcresTreated3.focus();
			}
		}	
		function displaySprayListData3(){
			for (i = 0; i < document.frmadd.SprayListID3.options.length; i++) {
			   if (document.frmadd.SprayListID3.options [i].selected == true) {
    				maxappvalues = document.frmadd.SprayListID3.options [i].value.split("|");
					document.frmadd.MaxAppUse3.value = maxappvalues[1];
					document.frmadd.MaxAppSeason3.value = maxappvalues[2];
					document.frmadd.Units3.value = maxappvalues[3];
  				 }
			}
		}	
		
		function calculateTotal2(){
			if (document.frmadd.RateAcre2.value == "")
				document.frmadd.RateAcre2.value = 0
			if (document.frmadd.AcresTreated2.value == "")
				document.frmadd.AcresTreated2.value = 0
			if (document.frmadd.TotalMaterialApplied2.value == "")
				document.frmadd.TotalMaterialApplied2.value = 0
				
			if (!(isNaN(document.frmadd.AcresTreated2.value))) {
				if (!(isNaN(document.frmadd.RateAcre2.value)) && (document.frmadd.RateAcre2.value != 0)) {
					document.frmadd.TotalMaterialApplied2.value = (Math.round(100*document.frmadd.AcresTreated2.value *  document.frmadd.RateAcre2.value)/100) ;
				}
				else{
					if (!(isNaN(document.frmadd.TotalMaterialApplied2.value)) && document.frmadd.TotalMaterialApplied2.value != 0) {
						document.frmadd.RateAcre2.value = (Math.round(100*document.frmadd.TotalMaterialApplied2.value / document.frmadd.AcresTreated2.value)/100) ;
					}
					else
					{
						//alert("Please enter a numeric value.");
						document.frmadd.RateAcre2.value = 0;
						document.frmadd.RateAcre2.focus();
					}
				}
			}
			else{
				//alert("Please enter a numeric value.");
				document.frmadd.AcresTreated2.value = 0;
				document.frmadd.AcresTreated2.focus();
			}
		}	
		function displaySprayListData2(){
			for (i = 0; i < document.frmadd.SprayListID2.options.length; i++) {
			   if (document.frmadd.SprayListID2.options [i].selected == true) {
    				maxappvalues = document.frmadd.SprayListID2.options [i].value.split("|");
					document.frmadd.MaxAppUse2.value = maxappvalues[1];
					document.frmadd.MaxAppSeason2.value = maxappvalues[2];
					document.frmadd.Units2.value = maxappvalues[3];
  				 }
			}
		}	
		
		function calculateTotal1(){
			if (document.frmadd.RateAcre1.value == "")
				document.frmadd.RateAcre1.value = 0
			if (document.frmadd.AcresTreated1.value == "")
				document.frmadd.AcresTreated1.value = 0
			if (document.frmadd.TotalMaterialApplied1.value == "")
				document.frmadd.TotalMaterialApplied1.value = 0
				
			if (!(isNaN(document.frmadd.AcresTreated1.value))) {
				if (!(isNaN(document.frmadd.RateAcre1.value)) && (document.frmadd.RateAcre1.value != 0)) {
					document.frmadd.TotalMaterialApplied1.value = (Math.round(100*document.frmadd.AcresTreated1.value *  document.frmadd.RateAcre1.value)/100) ;
				}
				else{
					if (!(isNaN(document.frmadd.TotalMaterialApplied1.value)) && document.frmadd.TotalMaterialApplied1.value != 0) {
						document.frmadd.RateAcre1.value = (Math.round(100*document.frmadd.TotalMaterialApplied1.value / document.frmadd.AcresTreated1.value)/100) ;
					}
					else
					{
						////alert("Please enter a numeric value.");
						document.frmadd.RateAcre1.value = 0;
						document.frmadd.RateAcre1.focus();
					}
				}

			}
			else{
				//alert("Please enter a numeric value.");
				document.frmadd.AcresTreated1.value = 0;
				document.frmadd.AcresTreated1.focus();
			}
		}	
		function displaySprayListData1(){
			for (i = 0; i < document.frmadd.SprayListID1.options.length; i++) {
			   if (document.frmadd.SprayListID1.options [i].selected == true) {
    				maxappvalues = document.frmadd.SprayListID1.options [i].value.split("|");
					document.frmadd.MaxAppUse1.value = maxappvalues[1];
					document.frmadd.MaxAppSeason1.value = maxappvalues[2];
					document.frmadd.Units1.value = maxappvalues[3];
  				 }
			}
		}	

	</script>	
  