<?php
if (session_id() == "") session_start(); // Initialize Session data
ob_start(); // Turn on output buffering
?>
<?php include_once "ewcfg12.php" ?>
<?php
if(EW_USE_ADODB) {
    include_once "adodb5/adodb.inc.php";
} else {
    include_once "ewmysql12.php";
}
?>
<?php include_once "phpfn12.php" ?>
<?php include_once "a_customersinfo.php" ?>
<?php include_once "usersinfo.php" ?>
<?php include_once "a_salesgridcls.php" ?>
<?php include_once "userfn12.php" ?>
<?php include_once "function_1.php"?>
<?php include_once "load_1.php"?>
<?php include_once "cod_html3.php"?>
<?php include_once "cod_html4.php"?>
<?php include_once "function_2.php"?>
<?php include_once "funct_1.php"?>
<?php include_once "funct_2.php"?>
<?php include_once "cod_html_7.php"?>
<?php include_once "html_if.php"?>
<?php

//
// Page class
//

$a_customers_edit = NULL; // Initialize page object first

class ca_customers_edit extends ca_customers {

	// Page ID
	var $PageID = 'edit';

	// Project ID
	var $ProjectID = "{B36B93AF-B58F-461B-B767-5F08C12493E9}";

	// Table name
	var $TableName = 'a_customers';

	// Page object name
	var $PageObjName = 'a_customers_edit';


	var $PageHeader;
	var $PageFooter;


	var $Token = "";
	var $TokenTimeout = 0;
	var $CheckToken = EW_CHECK_TOKEN;
	var $CheckTokenFn = "ew_CheckToken";
	var $CreateTokenFn = "ew_CreateToken";

	//
	// Page class constructor
	//
	function __construct() {
            session_start();

		$_SESSION['Page'] = &$this;
		$this->TokenTimeout = ew_SessionTimeoutTime();

		// Language object
                $_SESSION['Language']  = new cLanguage();

		// Parent constuctor
		parent::__construct();

		// Table object (a_customers)
		if (!isset($_SESSION['a_customers']) || get_class($_SESSION['a_customers']) == "ca_customers") {
			$_SESSION['a_customers'] = &$this;
			$_SESSION['Table'] = &$_SESSION['a_customers'];
		}

		// Table object (users)
		if (!isset($_SESSION['users'])) $_SESSION['users'] = new cusers();

		// Page ID
		if (!defined("EW_PAGE_ID"))
			define("EW_PAGE_ID", 'edit', TRUE);

		// Table name (for backward compatibility)
		if (!defined("EW_TABLE_NAME"))
			define("EW_TABLE_NAME", 'a_customers', TRUE);

		// Start timer
		if (!isset($GLOBALS["gTimer"])) $_SESSION['gTimer'] = new cTimer();

		// Open connection
        $_SESSION['conn'] = ew_Connect($this->DBID);


        // User table object (users)
        $_SESSION['UserTable']  = new cusers();
        $_SESSION['UserTableConn']  = Conn($_SESSION['UserTable']->DBID);
	}

	// 
	//  Page_Init
	//
	function Page_Init() {
		$init= new funct_1();
		$init->Page_Init1($this);
	}

	//
	// Page_Terminate
	//
	function Page_Terminate($url = "") {
        $term= new funct_1();
        $term->Page_Terminate1($url, $this);
	}
	var $FormClassName = "form-horizontal ewForm ewEditForm";
	var $DbMasterFilter;
	var $DbDetailFilter;
	var $DisplayRecs = 1;
	var $StartRec;
	var $StopRec;
	var $TotalRecs = 0;
	var $RecRange = 10;
	var $Pager;
	var $RecCnt;
	var $RecKey = array();
	var $Recordset;

	// 
	// Page main
	//
	function Page_Main() {
           session_start();
		
                $_SESSION['Language']  = new cLanguage();

		// Load current record
		$bLoadCurrentRecord = FALSE;

		$bMatchRecord = FALSE;

		// Load key from QueryString
		if ($_GET["Customer_ID"] <> "") {
			$this->Customer_ID->setQueryStringValue($_GET["Customer_ID"]);
			$this->RecKey["Customer_ID"] = $this->Customer_ID->QueryStringValue;
		} else {
			$bLoadCurrentRecord = TRUE;
		}

		// Set up Breadcrumb
		$this->SetupBreadcrumb();

		// Load recordset
		$this->StartRec = 1; // Initialize start position
		if ($this->Recordset = $this->LoadRecordset()) // Load records
			$this->TotalRecs = $this->Recordset->RecordCount(); // Get record count
		if ($this->TotalRecs <= 0) { // No record found
			if ($this->getSuccessMessage() == "" && $this->getFailureMessage() == "")
				$this->setFailureMessage($_SESSION['Language']->Phrase("NoRecord")); // Set no record message
			$this->Page_Terminate("a_customerslist.php"); // Return to list page
		} elseif ($bLoadCurrentRecord) { // Load current record position
			$this->SetUpStartRec(); // Set up start record position

			// Point to current record
			if (intval($this->StartRec) <= intval($this->TotalRecs)) {
				$bMatchRecord = TRUE;
				$this->Recordset->Move($this->StartRec-1);
			}
		} else { // Match key values
			while (!$this->Recordset->EOF) {
				if (strval($this->Customer_ID->CurrentValue) == strval($this->Recordset->fields('Customer_ID'))) {
					$this->setStartRecordNumber($this->StartRec); // Save record position
					$bMatchRecord = TRUE;
					break;
				} else {
					$this->StartRec++;
					$this->Recordset->MoveNext();
				}
			}
		}

		// Process form if post back
		if ($_POST["a_edit"] <> "") {
			$this->CurrentAction = $_POST["a_edit"]; // Get action code
			$this->LoadFormValues(); // Get form values

			// Set up detail parameters
			$this->SetUpDetailParms();
		} else {
			$this->CurrentAction = "I"; // Default action is display
		}

		// Validate form if post back
		if ($_POST["a_edit"] <> "") {
			if (!$this->ValidateForm()) {
				$this->CurrentAction = ""; // Form error, reset action
				$this->setFailureMessage($_SESSION['gsFormError']);
				$this->EventCancelled = TRUE; // Event cancelled
				$this->RestoreFormValues();
			}
		}
		switch ($this->CurrentAction) {
			case "I": // Get a record to display
				if (!$bMatchRecord) {
					if ($this->getSuccessMessage() == "" && $this->getFailureMessage() == "")
						$this->setFailureMessage($_SESSION['Language']->Phrase("NoRecord")); // Set no record message
					$this->Page_Terminate("a_customerslist.php"); // Return to list page
				} else {
					$this->LoadRowValues($this->Recordset); // Load row values
				}

				// Set up detail parameters
				$this->SetUpDetailParms();
				break;
			Case "U": // Update
				if ($this->getCurrentDetailTable() <> "") // Master/detail edit
					$sReturnUrl = $this->GetViewUrl(EW_TABLE_SHOW_DETAIL . "=" . $this->getCurrentDetailTable()); // Master/Detail view page
				else
					$sReturnUrl = $this->getReturnUrl();
				if (ew_GetPageName($sReturnUrl) == "a_customerslist.php")
					$sReturnUrl = $this->AddMasterUrl($this->GetListUrl()); // List page, return to list page with correct master key if necessary
				$this->SendEmail = TRUE; // Send email on update success
				if ($this->EditRow()) { // Update record based on key

					// Begin of modification Disable Add/Edit Success Message Box, by Masino Sinaga, August 1, 2012
					if (MS_SHOW_EDIT_SUCCESS_MESSAGE==TRUE) {
						if ($this->getSuccessMessage() == "")
							$this->setSuccessMessage($_SESSION['Language']->Phrase("UpdateSuccess")); // Update success
					}

					// Begin of modification Disable Add/Edit Success Message Box, by Masino Sinaga, August 1, 2012
					$this->Page_Terminate($sReturnUrl); // Return to caller
				} elseif ($this->getFailureMessage() == $_SESSION['Language']->Phrase("NoRecord")) {
					$this->Page_Terminate($sReturnUrl); // Return to caller
				} else {
					$this->EventCancelled = TRUE; // Event cancelled
					$this->RestoreFormValues(); // Restore form values if update failed

					// Set up detail parameters
					$this->SetUpDetailParms();
				}break;
            default:
                echo 'it is not present in any case';
		}

		// Render the record
		$this->RowType = EW_ROWTYPE_EDIT; // Render as Edit
		$this->ResetAttrs();
		$this->RenderRow();
	}

	// Set up starting record parameters
	function SetUpStartRec() {
		$start= new funct_2();
		$start->SetUpStartRec1($this);
	}

	// Get upload files
	function GetUploadFiles() {
		

		// Get upload data
	}

	// Load form values
	function LoadFormValues() {

		// Load from form
		session_start();
		if (!$this->Customer_ID->FldIsDetailKey)
			$this->Customer_ID->setFormValue($_SESSION['objForm']->GetValue("x_Customer_ID"));
		if (!$this->Customer_Number->FldIsDetailKey) {
			$this->Customer_Number->setFormValue($_SESSION['objForm']->GetValue("x_Customer_Number"));
		}
		if (!$this->Customer_Name->FldIsDetailKey) {
			$this->Customer_Name->setFormValue($_SESSION['objForm']->GetValue("x_Customer_Name"));
		}
		if (!$this->Address->FldIsDetailKey) {
			$this->Address->setFormValue($_SESSION['objForm']->GetValue("x_Address"));
		}
		if (!$this->City->FldIsDetailKey) {
			$this->City->setFormValue($_SESSION['objForm']->GetValue("x_City"));
		}
		if (!$this->Country->FldIsDetailKey) {
			$this->Country->setFormValue($_SESSION['objForm']->GetValue("x_Country"));
		}
		if (!$this->Contact_Person->FldIsDetailKey) {
			$this->Contact_Person->setFormValue($_SESSION['objForm']->GetValue("x_Contact_Person"));
		}
		if (!$this->Phone_Number->FldIsDetailKey) {
			$this->Phone_Number->setFormValue($_SESSION['objForm']->GetValue("x_Phone_Number"));
		}
		if (!$this->_Email->FldIsDetailKey) {
			$this->_Email->setFormValue($_SESSION['objForm']->GetValue("x__Email"));
		}
		if (!$this->Mobile_Number->FldIsDetailKey) {
			$this->Mobile_Number->setFormValue($_SESSION['objForm']->GetValue("x_Mobile_Number"));
		}
		if (!$this->Notes->FldIsDetailKey) {
			$this->Notes->setFormValue($_SESSION['objForm']->GetValue("x_Notes"));
		}
		if (!$this->Balance->FldIsDetailKey) {
			$this->Balance->setFormValue($_SESSION['objForm']->GetValue("x_Balance"));
		}
		if (!$this->Date_Added->FldIsDetailKey) {
			$this->Date_Added->setFormValue($_SESSION['objForm']->GetValue("x_Date_Added"));
			$this->Date_Added->CurrentValue = ew_UnFormatDateTime($this->Date_Added->CurrentValue, 0);
		}
		if (!$this->Added_By->FldIsDetailKey) {
			$this->Added_By->setFormValue($_SESSION['objForm']->GetValue("x_Added_By"));
		}
		if (!$this->Date_Updated->FldIsDetailKey) {
			$this->Date_Updated->setFormValue($_SESSION['objForm']->GetValue("x_Date_Updated"));
			$this->Date_Updated->CurrentValue = ew_UnFormatDateTime($this->Date_Updated->CurrentValue, 0);
		}
		if (!$this->Updated_By->FldIsDetailKey) {
			$this->Updated_By->setFormValue($_SESSION['objForm']->GetValue("x_Updated_By"));
		}
	}

	// Restore form values
	function RestoreFormValues() {
		
		$this->LoadRow();
		$this->Customer_ID->CurrentValue = $this->Customer_ID->FormValue;
		$this->Customer_Number->CurrentValue = $this->Customer_Number->FormValue;
		$this->Customer_Name->CurrentValue = $this->Customer_Name->FormValue;
		$this->Address->CurrentValue = $this->Address->FormValue;
		$this->City->CurrentValue = $this->City->FormValue;
		$this->Country->CurrentValue = $this->Country->FormValue;
		$this->Contact_Person->CurrentValue = $this->Contact_Person->FormValue;
		$this->Phone_Number->CurrentValue = $this->Phone_Number->FormValue;
		$this->_Email->CurrentValue = $this->_Email->FormValue;
		$this->Mobile_Number->CurrentValue = $this->Mobile_Number->FormValue;
		$this->Notes->CurrentValue = $this->Notes->FormValue;
		$this->Balance->CurrentValue = $this->Balance->FormValue;
		$this->Date_Added->CurrentValue = $this->Date_Added->FormValue;
		$this->Date_Added->CurrentValue = ew_UnFormatDateTime($this->Date_Added->CurrentValue, 0);
		$this->Added_By->CurrentValue = $this->Added_By->FormValue;
		$this->Date_Updated->CurrentValue = $this->Date_Updated->FormValue;
		$this->Date_Updated->CurrentValue = ew_UnFormatDateTime($this->Date_Updated->CurrentValue, 0);
		$this->Updated_By->CurrentValue = $this->Updated_By->FormValue;
	}


	// Update record based on key values
	function EditRow() {
            session_start();
		$_SESSION['Language']  = new cLanguage();
		$sFilter = $this->KeyFilter();
		$sFilter = $this->ApplyUserIDFilters($sFilter);
		$conn = &$this->Connection();
		$this->CurrentFilter = $sFilter;
		$sSql = $this->SQL();
		$conn->raiseErrorFn = $_SESSION['EW_ERROR_FN']; // v11.0.4
		$rs = $conn->Execute($sSql);
		$conn->raiseErrorFn = '';
		if ($rs === FALSE)
			return FALSE;
		if ($rs->EOF) {
			$this->setFailureMessage($_SESSION['Language']->Phrase("NoRecord")); // Set no record message
			$EditRow = FALSE; // Update Failed
		} else {

			// Begin transaction
			if ($this->getCurrentDetailTable() <> "")
				$conn->BeginTrans();

			// Save old values
			$rsold = &$rs->fields;
			$this->LoadDbValues($rsold);
			$rsnew = array();

			// Customer_Number
			$this->Customer_Number->SetDbValueDef($rsnew, $this->Customer_Number->CurrentValue, "", $this->Customer_Number->ReadOnly);

			// Customer_Name
			$this->Customer_Name->SetDbValueDef($rsnew, $this->Customer_Name->CurrentValue, "", $this->Customer_Name->ReadOnly);

			// Address
			$this->Address->SetDbValueDef($rsnew, $this->Address->CurrentValue, "", $this->Address->ReadOnly);

			// City
			$this->City->SetDbValueDef($rsnew, $this->City->CurrentValue, "", $this->City->ReadOnly);

			// Country
			$this->Country->SetDbValueDef($rsnew, $this->Country->CurrentValue, "", $this->Country->ReadOnly);

			// Contact_Person
			$this->Contact_Person->SetDbValueDef($rsnew, $this->Contact_Person->CurrentValue, "", $this->Contact_Person->ReadOnly);

			// Phone_Number
			$this->Phone_Number->SetDbValueDef($rsnew, $this->Phone_Number->CurrentValue, "", $this->Phone_Number->ReadOnly);

			// Email
			$this->_Email->SetDbValueDef($rsnew, $this->_Email->CurrentValue, "", $this->_Email->ReadOnly);

			// Mobile_Number
			$this->Mobile_Number->SetDbValueDef($rsnew, $this->Mobile_Number->CurrentValue, "", $this->Mobile_Number->ReadOnly);

			// Notes
			$this->Notes->SetDbValueDef($rsnew, $this->Notes->CurrentValue, "", $this->Notes->ReadOnly);

			// Balance
			$this->Balance->SetDbValueDef($rsnew, $this->Balance->CurrentValue, NULL, $this->Balance->ReadOnly);

			// Date_Added
			$this->Date_Added->SetDbValueDef($rsnew, $this->Date_Added->CurrentValue, NULL, $this->Date_Added->ReadOnly);

			// Added_By
			$this->Added_By->SetDbValueDef($rsnew, $this->Added_By->CurrentValue, NULL, $this->Added_By->ReadOnly);

			// Date_Updated
			$this->Date_Updated->SetDbValueDef($rsnew, ew_CurrentDateTime(), NULL);
			$rsnew['Date_Updated'] = &$this->Date_Updated->DbValue;

			// Updated_By
			$this->Updated_By->SetDbValueDef($rsnew, CurrentUserName(), NULL);
			$rsnew['Updated_By'] = &$this->Updated_By->DbValue;

			// Call Row Updating event
			$bUpdateRow = $this->Row_Updating($rsold, $rsnew);
			if ($bUpdateRow) {
				$conn->raiseErrorFn = $_SESSION['EW_ERROR_FN']; // v11.0.4
				if (count($rsnew) > 0)
					$EditRow = $this->Update($rsnew);
				else
					$EditRow = TRUE; // No field to update
				$conn->raiseErrorFn = '';
				if ($EditRow) {
				}

				// Update detail records
				if ($EditRow) {
					$DetailTblVar = explode(",", $this->getCurrentDetailTable());
					if (in_array("a_sales", $DetailTblVar) && $_SESSION['a_sales']->DetailEdit) {
						if (!isset($_SESSION['a_sales_grid'])) $_SESSION['a_sales_grid'] = new ca_sales_grid(); // Get detail page object
						$EditRow = $_SESSION['a_sales_grid']->GridUpdate();
					}
				}

				// Commit/Rollback transaction
				if ($this->getCurrentDetailTable() <> "") {
					if ($EditRow) {
						$conn->CommitTrans(); // Commit transaction
					} else {
						$conn->RollbackTrans(); // Rollback transaction
					}
				}
			} else {
				if ($this->getSuccessMessage() <> "" || $this->getFailureMessage() <> "") {

					// Use the message, do nothing
				} elseif ($this->CancelMessage <> "") {
					$this->setFailureMessage($this->CancelMessage);
					$this->CancelMessage = "";
				} else {
					$this->setFailureMessage($_SESSION['Language']->Phrase("UpdateCancelled"));
				}
				$EditRow = FALSE;
			}
		}

		// Call Row_Updated event
		if ($EditRow)
			$this->Row_Updated();
		$rs->Close();
		return $EditRow;
	}

	// Build export filter for selected records
	function BuildExportSelectedFilter() {
		
		$sWrkFilter = "";
		if ($this->Export <> "") {
			$sWrkFilter = $this->GetKeyFilter();
		}
		return $sWrkFilter;
	}

	// Set up detail parms based on QueryString
	function SetUpDetailParms() {
            session_start();

		// Get the keys for master table
		if (isset($_GET[EW_TABLE_SHOW_DETAIL])) {
			$sDetailTblVar = $_GET[EW_TABLE_SHOW_DETAIL];
			$this->setCurrentDetailTable($sDetailTblVar);
		} else {
			$sDetailTblVar = $this->getCurrentDetailTable();
		}
		if ($sDetailTblVar <> "") {
			$DetailTblVar = explode(",", $sDetailTblVar);
			if (in_array("a_sales", $DetailTblVar)) {
				if (!isset($_SESSION['a_sales_grid']))
					$_SESSION['a_sales_grid'] = new ca_sales_grid;
				if ($_SESSION['a_sales_grid']->DetailEdit) {
					$_SESSION['a_sales_grid']->CurrentMode = "edit";
					$_SESSION['a_sales_grid']->CurrentAction = "gridedit";

					// Save current master table to detail table
					$_SESSION['a_sales_grid']->setCurrentMasterTable($this->TableVar);
					$_SESSION['a_sales_grid']->setStartRecordNumber(1);
					$_SESSION['a_sales_grid']->Customer_ID->FldIsDetailKey = TRUE;
					$_SESSION['a_sales_grid']->Customer_ID->CurrentValue = $this->Customer_Number->CurrentValue;
					$_SESSION['a_sales_grid']->Customer_ID->setSessionValue($_SESSION['a_sales_grid']->Customer_ID->CurrentValue);
				}
			}
		}
	}

	// Set up Breadcrumb
	function SetupBreadcrumb() {
		
		$_SESSION['Breadcrumb'] = new cBreadcrumb();
		$url = substr(ew_CurrentUrl(), strrpos(ew_CurrentUrl(), "/")+1); // v11.0.4
		$_SESSION['Breadcrumb']->Add("list", $this->TableVar, $this->AddMasterUrl("a_customerslist.php"), "", $this->TableVar, TRUE);
		$PageId = "edit";
		$_SESSION['Breadcrumb']->Add("edit", $PageId, $url); // v11.0.4
	}

	// Page Load event
	function Page_Load() {

		//echo "Page Load";
	}

	// Page Unload event
	function Page_Unload() {

		//echo "Page Unload";
	}

	// Page Redirecting event
	function Page_Redirecting() {

		// Example:
		//$url = "your URL";

	}

	// Message Showing event
	// $type = ''|'success'|'failure'|'warning'
	function Message_Showing(&$msg, $type) {
		if ($type == 'success') {

			$msg = $msg."";
		} elseif ($type == 'failure') {

            $msg = $msg."";
		} elseif ($type == 'warning') {

            $msg = $msg."";
		}
	}

	// Page Render event
	function Page_Render() {

		//echo "Page Render";
	}

	// Page Data Rendering event
	function Page_DataRendering() {

		// Example:
		//$header = "your header";

	}

	// Page Data Rendered event
	function Page_DataRendered() {

		// Example:
		//$footer = "your footer";

	}

	// Form Custom Validate event
	function Form_CustomValidate() {

		// Return error message in CustomError
		return TRUE;
	}
}
?>
<?php ew_Header(FALSE) ?>
<?php

// Create page object
if (!isset($a_customers_edit)) $a_customers_edit = new ca_customers_edit();

// Page init
$a_customers_edit->Page_Init();

// Page main
$a_customers_edit->Page_Main();

// Begin of modification Displaying Breadcrumb Links in All Pages, by Masino Sinaga, May 4, 2012
getCurrentPageTitle(ew_CurrentPage());

// End of modification Displaying Breadcrumb Links in All Pages, by Masino Sinaga, May 4, 2012
// Global Page Rendering event (in userfn*.php)

Page_Rendering();

// Global auto switch table width style (in userfn*.php), by Masino Sinaga, January 7, 2015
AutoSwitchTableWidthStyle();

// Page Rendering event
$a_customers_edit->Page_Render();
?>
<?php include_once "header.php" ?>
<script type="text/javascript">

// Form object
var CurrentPageID = EW_PAGE_ID = "edit";
var CurrentForm = fa_customersedit = new ew_Form("fa_customersedit", "edit");



// Form_CustomValidate event
fa_customersedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!

 	// Your custom validation code here, return false if invalid. 
 	return true;
 }

// Use JavaScript validation or not
<?php if (EW_CLIENT_VALIDATE) { ?>
fa_customersedit.ValidateRequired = true;
<?php } else { ?>
fa_customersedit.ValidateRequired = false; 
<?php } ?>

// Dynamic selection lists
// Form object for search

</script>
<script type="text/javascript">

// Write your client script here, no need to add script tags.
</script>
<div class="ewToolbar">
<?php if (MS_SHOW_PHPMAKER_BREADCRUMBLINKS) { ?>
<?php $Breadcrumb->Render(); ?>
<?php } ?>
<?php if (MS_SHOW_MASINO_BREADCRUMBLINKS) { ?>
<?php echo htmlspecialchars(MasinoBreadcrumbLinks()); ?>
<?php } ?>
<?php if (MS_LANGUAGE_SELECTOR_VISIBILITY=="belowheader") { ?>
<?php echo $Language->SelectionForm(); ?>
<?php } ?>
<div class="clearfix"></div>
</div>
<?php $a_customers_edit->ShowPageHeader(); ?>
<?php
$a_customers_edit->ShowMessage();
?>
<?php // Begin of modification Customize Navigation/Pager Panel, by Masino Sinaga, May 2, 2012 ?>
<?php if ( (MS_PAGINATION_POSITION==1) || (MS_PAGINATION_POSITION==3) ) { ?>
<form name="ewPagerForm" class="form-horizontal ewForm ewPagerForm" action="<?php echo ew_CurrentPage() ?>">
	<?php if (MS_PAGINATION_STYLE==1) { // link ?>
        <?php
        $html2= new cod_html3();
        echo $html2->HTML4($a_customers_edit, $Language);
        ?>
	<?php } elseif (MS_PAGINATION_STYLE==2) { // button ?>
		<?php
        $html07= new cod_html_7();
        echo $html07->html_07($a_customers_edit, $Language);
        ?>
	<?php } // end of link or button ?>	
<div class="clearfix"></div>
</form>
<?php } ?>
<?php // End of modification Customize Navigation/Pager Panel, by Masino Sinaga, May 2, 2012 ?>
<form name="fa_customersedit" id="fa_customersedit" class="<?php echo $a_customers_edit->FormClassName ?>" action="<?php echo ew_CurrentPage() ?>" method="post">
<?php if ($a_customers_edit->CheckToken) { ?>
<input type="hidden" name="<?php echo EW_TOKEN_NAME ?>" value="<?php echo $a_customers_edit->Token ?>">
<?php } ?>
<input type="hidden" name="t" value="a_customers">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<div>
<?php if ($a_customers->Customer_ID->Visible) { // Customer_ID ?>
	<div id="r_Customer_ID" class="form-group">
		<label id="elh_a_customers_Customer_ID" class="col-sm-4 control-label ewLabel"><?php echo $a_customers->Customer_ID->FldCaption() ?></label>
		<div class="col-sm-8"><div<?php echo $a_customers->Customer_ID->CellAttributes() ?>>
<span id="el_a_customers_Customer_ID">
<span<?php echo $a_customers->Customer_ID->ViewAttributes() ?>>
<p class="form-control-static"><?php echo $a_customers->Customer_ID->EditValue ?></p></span>
</span>
<input type="hidden" data-table="a_customers" data-field="x_Customer_ID" name="x_Customer_ID" id="x_Customer_ID" value="<?php echo ew_HtmlEncode($a_customers->Customer_ID->CurrentValue) ?>">
<?php echo $a_customers->Customer_ID->CustomMsg ?></div></div>
	</div>
<?php } ?>
<?php if ($a_customers->Customer_Number->Visible) { // Customer_Number ?>
    <?php
    $html_if = new html_if();
    $html_if->html_code($a_customers, $Language);
    ?>
    <?php if ($a_customers->Balance->Visible) { // Balance ?>
        <div id="r_Balance" class="form-group">
            <label id="elh_a_customers_Balance" for="x_Balance"
                   class="col-sm-4 control-label ewLabel"><?php echo $a_customers->Balance->FldCaption() ?></label>
            <div class="col-sm-8">
                <div<?php echo $a_customers->Balance->CellAttributes() ?>>
<span id="el_a_customers_Balance">
<input type="text" data-table="a_customers" data-field="x_Balance" name="x_Balance" id="x_Balance" size="30"
       placeholder="<?php echo ew_HtmlEncode($a_customers->Balance->getPlaceHolder()) ?>"
       value="<?php echo $a_customers->Balance->EditValue ?>"<?php echo $a_customers->Balance->EditAttributes() ?>>
    <?php if (!$a_customers->Balance->ReadOnly && !$a_customers->Balance->Disabled && $a_customers->Balance->EditAttrs["readonly"] == "" && $a_customers->Balance->EditAttrs["disabled"] == "") { ?>
        <script type="text/javascript">
$('#x_Balance').autoNumeric('init', {aSep: ',', aDec: '.', mDec: '2', aForm: false});
</script>
    <?php } ?>
</span>
                    <?php echo $a_customers->Balance->CustomMsg ?></div>
            </div>
        </div>
    <?php } ?>
</div>
    <span id="el_a_customers_Date_Added">
<input type="hidden" data-table="a_customers" data-field="x_Date_Added" name="x_Date_Added" id="x_Date_Added"
       value="<?php echo ew_HtmlEncode($a_customers->Date_Added->CurrentValue) ?>">
</span>
    <span id="el_a_customers_Added_By">
<input type="hidden" data-table="a_customers" data-field="x_Added_By" name="x_Added_By" id="x_Added_By"
       value="<?php echo ew_HtmlEncode($a_customers->Added_By->CurrentValue) ?>">
</span>
    <?php
    if (in_array("a_sales", explode(",", $a_customers->getCurrentDetailTable())) && $a_sales->DetailEdit) {
        ?>
        <?php if ($a_customers->getCurrentDetailTable() <> "") { ?>
            <h4 class="ewDetailCaption"><?php echo $Language->TablePhrase("a_sales", "TblCaption") ?></h4>
        <?php } ?>
        <?php include_once "a_salesgrid.php" ?>
    <?php } ?>
    <div class="form-group">
        <div class="col-sm-offset-4 col-sm-8">
            <button class="btn btn-primary ewButton" name="btnAction" id="btnAction"
                    type="submit"><?php echo $Language->Phrase("SaveBtn") ?></button>
            <button class="btn btn-danger ewButton" name="btnCancel" id="btnCancel" type="button"
                    data-href="<?php echo $a_customers_edit->getReturnUrl() ?>"><?php echo $Language->Phrase("CancelBtn") ?></button>
        </div>
    </div>
    <?php // Begin of modification Customize Navigation/Pager Panel, by Masino Sinaga, May 2, 2012 ?>
    <?php if ((MS_PAGINATION_POSITION == 2) || (MS_PAGINATION_POSITION == 3)) { ?>
        <?php if (MS_PAGINATION_STYLE == 1) { // link ?>
            <?php
            $html = new cod_html3();
            echo $html->HTML4($a_customers_edit, $Language);
            ?>

        <?php } elseif (MS_PAGINATION_STYLE == 2) { // button ?>
            <?php if (!isset($a_customers_edit->Pager)) $a_customers_edit->Pager = new cPrevNextPager($a_customers_edit->StartRec, $a_customers_edit->DisplayRecs, $a_customers_edit->TotalRecs) ?>
            <?php if ($a_customers_edit->Pager->RecordCount > 0) { ?>
                <?php if (($a_customers_edit->Pager->PageCount == 1) && ($a_customers_edit->Pager->CurrentPage == 1) && (MS_SHOW_PAGENUM_IF_REC_NOT_OVER_PAGESIZE == FALSE)) { ?>
                <?php } else { // end MS_SHOW_PAGENUM_IF_REC_NOT_OVER_PAGESIZE==FALSE ?>
                    <div class="ewPager">
                        <span><?php echo $Language->Phrase("Page") ?>&nbsp;</span>
                        <div class="ewPrevNext">
                            <div class="input-group">
                                <div class="input-group-btn">
                                    <!--first page button-->
                                    <?php if ($a_customers_edit->Pager->FirstButton->Enabled) { ?>
                                        <?php if ($Language->Phrase("dir") == "rtl") { // begin of rtl ?>
                                            <a class="btn btn-default btn-sm"
                                               title="<?php echo $Language->Phrase("PagerFirst") ?>"
                                               href="<?php echo $a_customers_edit->PageUrl() ?>start=<?php echo $a_customers_edit->Pager->FirstButton->Start ?>"><span
                                                        class="icon-last ewIcon"></span></a>
                                        <?php } else { // else of rtl ?>
                                            <a class="btn btn-default btn-sm"
                                               title="<?php echo $Language->Phrase("PagerFirst") ?>"
                                               href="<?php echo $a_customers_edit->PageUrl() ?>start=<?php echo $a_customers_edit->Pager->FirstButton->Start ?>"><span
                                                        class="icon-first ewIcon"></span></a>
                                        <?php } // end of rtl ?>
                                    <?php } else { ?>
                                        <?php if ($Language->Phrase("dir") == "rtl") { // begin of rtl ?>
                                            <a class="btn btn-default btn-sm disabled"
                                               title="<?php echo $Language->Phrase("PagerFirst") ?>"><span
                                                        class="icon-last ewIcon"></span></a>
                                        <?php } else { // else of rtl ?>
                                            <a class="btn btn-default btn-sm disabled"
                                               title="<?php echo $Language->Phrase("PagerFirst") ?>"><span
                                                        class="icon-first ewIcon"></span></a>
                                        <?php } // end of rtl ?>
                                    <?php } ?>
                                    <!--previous page button-->
                                    <?php if ($a_customers_edit->Pager->PrevButton->Enabled) { ?>
                                        <?php if ($Language->Phrase("dir") == "rtl") { // begin of rtl ?>
                                            <a class="btn btn-default btn-sm"
                                               title="<?php echo $Language->Phrase("PagerPrevious") ?>"
                                               href="<?php echo $a_customers_edit->PageUrl() ?>start=<?php echo $a_customers_edit->Pager->PrevButton->Start ?>"><span
                                                        class="icon-next ewIcon"></span></a>
                                        <?php } else { // else of rtl ?>
                                            <a class="btn btn-default btn-sm"
                                               title="<?php echo $Language->Phrase("PagerPrevious") ?>"
                                               href="<?php echo $a_customers_edit->PageUrl() ?>start=<?php echo $a_customers_edit->Pager->PrevButton->Start ?>"><span
                                                        class="icon-prev ewIcon"></span></a>
                                        <?php } // end of rtl ?>
                                    <?php } else { ?>
                                        <?php if ($Language->Phrase("dir") == "rtl") { // begin of rtl ?>
                                            <a class="btn btn-default btn-sm disabled"
                                               title="<?php echo $Language->Phrase("PagerPrevious") ?>"><span
                                                        class="icon-next ewIcon"></span></a>
                                        <?php } else { // else of rtl ?>
                                            <a class="btn btn-default btn-sm disabled"
                                               title="<?php echo $Language->Phrase("PagerPrevious") ?>"><span
                                                        class="icon-prev ewIcon"></span></a>
                                        <?php } // end of rtl ?>
                                    <?php } ?>
                                </div>
                                <!--current page number-->
                                <input class="form-control input-sm" type="text" name="<?php echo EW_TABLE_PAGE_NO ?>"
                                       value="<?php echo $a_customers_edit->Pager->CurrentPage ?>">
                                <div class="input-group-btn">
                                    <!--next page button-->
                                    <?php if ($a_customers_edit->Pager->NextButton->Enabled) { ?>
                                        <?php if ($Language->Phrase("dir") == "rtl") { // begin of rtl ?>
                                            <a class="btn btn-default btn-sm"
                                               title="<?php echo $Language->Phrase("PagerNext") ?>"
                                               href="<?php echo $a_customers_edit->PageUrl() ?>start=<?php echo $a_customers_edit->Pager->NextButton->Start ?>"><span
                                                        class="icon-prev ewIcon"></span></a>
                                        <?php } else { // else of rtl ?>
                                            <a class="btn btn-default btn-sm"
                                               title="<?php echo $Language->Phrase("PagerNext") ?>"
                                               href="<?php echo $a_customers_edit->PageUrl() ?>start=<?php echo $a_customers_edit->Pager->NextButton->Start ?>"><span
                                                        class="icon-next ewIcon"></span></a>
                                        <?php } // end of rtl ?>
                                    <?php } else { ?>
                                        <?php if ($Language->Phrase("dir") == "rtl") { // begin of rtl ?>
                                            <a class="btn btn-default btn-sm disabled"
                                               title="<?php echo $Language->Phrase("PagerNext") ?>"><span
                                                        class="icon-prev ewIcon"></span></a>
                                        <?php } else { // else of rtl ?>
                                            <a class="btn btn-default btn-sm disabled"
                                               title="<?php echo $Language->Phrase("PagerNext") ?>"><span
                                                        class="icon-next ewIcon"></span></a>
                                        <?php } // end of rtl ?>
                                    <?php } ?>
                                    <!--last page button-->
                                    <?php if ($a_customers_edit->Pager->LastButton->Enabled) { ?>
                                        <?php if ($Language->Phrase("dir") == "rtl") { // begin of rtl ?>
                                            <a class="btn btn-default btn-sm"
                                               title="<?php echo $Language->Phrase("PagerLast") ?>"
                                               href="<?php echo $a_customers_edit->PageUrl() ?>start=<?php echo $a_customers_edit->Pager->LastButton->Start ?>"><span
                                                        class="icon-first ewIcon"></span></a>
                                        <?php } else { // else of rtl ?>
                                            <a class="btn btn-default btn-sm"
                                               title="<?php echo $Language->Phrase("PagerLast") ?>"
                                               href="<?php echo $a_customers_edit->PageUrl() ?>start=<?php echo $a_customers_edit->Pager->LastButton->Start ?>"><span
                                                        class="icon-last ewIcon"></span></a>
                                        <?php } // end of rtl ?>
                                    <?php } else { ?>
                                        <?php if ($Language->Phrase("dir") == "rtl") { // begin of rtl ?>
                                            <a class="btn btn-default btn-sm disabled"
                                               title="<?php echo $Language->Phrase("PagerLast") ?>"><span
                                                        class="icon-first ewIcon"></span></a>
                                        <?php } else { // else of rtl ?>
                                            <a class="btn btn-default btn-sm disabled"
                                               title="<?php echo $Language->Phrase("PagerLast") ?>"><span
                                                        class="icon-last ewIcon"></span></a>
                                        <?php } // end of rtl ?>
                                    <?php } ?>
                                </div>
                            </div>
                        </div>
                        <span>&nbsp;<?php echo $Language->Phrase("of") ?>
                            &nbsp;<?php echo $a_customers_edit->Pager->PageCount ?></span>
                    </div>
                <?php } // end MS_SHOW_PAGENUM_IF_REC_NOT_OVER_PAGESIZE==FALSE ?>
            <?php } ?>
        <?php } // end of link or button ?>

        <div class="clearfix"></div>
    <?php } ?>
    <?php // End of modification Customize Navigation/Pager Panel, by Masino Sinaga, May 2, 2012 ?>
</form>
    <script type="text/javascript">
        fa_customersedit.Init();
    </script>
<?php
$a_customers_edit->ShowPageFooter();
if (EW_DEBUG_ENABLED)
    echo ew_DebugMsg();
?>
    <script type="text/javascript">

        // Write your table-specific startup script here
        // document.write("page loaded");

    </script>
<?php if (MS_ENTER_MOVING_CURSOR_TO_NEXT_FIELD) { ?>
    <script type="text/javascript">
        $(document).ready(function () {
            $("#fa_customersedit:first *:input[type!=hidden]:first").focus(), $("input").keydown(function (i) {
                if (13 == i.which) {
                    var e = $(this).closest("form").find(":input:visible:enabled"), n = e.index(this);
                    n == e.length - 1 || (e.eq(e.index(this) + 1).focus(), i.preventDefault())
                } else 113 == i.which && $("#btnAction").click()
            }), $("select").keydown(function (i) {
                if (13 == i.which) {
                    var e = $(this).closest("form").find(":input:visible:enabled"), n = e.index(this);
                    n == e.length - 1 || (e.eq(e.index(this) + 1).focus(), i.preventDefault())
                } else 113 == i.which && $("#btnAction").click()
            }), $("radio").keydown(function (i) {
                if (13 == i.which) {
                    var e = $(this).closest("form").find(":input:visible:enabled"), n = e.index(this);
                    n == e.length - 1 || (e.eq(e.index(this) + 1).focus(), i.preventDefault())
                } else 113 == i.which && $("#btnAction").click()
            })
        });
    </script>
<?php } ?>
<?php if ($a_customers->Export == "") { ?>
    
    
    <script type="text/javascript">
        $('#btnAction').attr('onclick', 'return alertifyEdit(this)');

        function alertifyEdit(obj) { <?php global $Language; ?> if (fa_customersedit.Validate() == true) {
            alertify.confirm("<?php echo $Language->Phrase('AlertifyEditConfirm'); ?>", function (e) {
                if (e) {
                    $(window).unbind('beforeunload');
                    alertify.success("<?php echo $Language->Phrase('AlertifyEdit'); ?>");
                    $("#fa_customersedit").submit();
                }
            }).set("title", "<?php echo $Language->Phrase('AlertifyConfirm'); ?>").set("defaultFocus", "cancel").set('oncancel', function (closeEvent) {
                alertify.error('<?php echo $Language->Phrase('AlertifyCancel'); ?>');
            }).set('labels', {
                ok: '<?php echo $Language->Phrase("MyOKMessage"); ?>!',
                cancel: '<?php echo $Language->Phrase("MyCancelMessage"); ?>'
            });
        }
            return false;
        }
    </script>
<?php } ?>
<?php include_once "footer.php" ?>
<?php
$a_customers_edit->Page_Terminate();
}?>
