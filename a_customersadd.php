<?php
if (session_id() == "") session_start(); // initialize Session data
ob_start(); // turn on output buffering
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
<?php include_once "cod_html4.php"?>
<?php include_once "function_2.php"?>
<?php include_once "funct_1.php"?>
<?php include_once "html_if.php"?>

<?php

//
// page class
//

$a_customers_add = NULL; // Initialize page object first

class ca_customers_add extends ca_customers {


	// page id

	// Page ID

	var $PageID = 'add';

	// project id
	var $ProjectID = "{B36B93AF-B58F-461B-B767-5F08C12493E9}";

	// Table name
	var $TableName = 'a_customers';

	// Page object name
	var $PageObjName = 'a_customers_add';


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
		if (!isset($_SESSION['a_customers']) || get_class($_SESSION['a_customers']) == 'ca_customers') {
			$_SESSION['a_customers'] = &$this;
			$_SESSION['Table'] = &$_SESSION['a_customers'];
		}

		// Table object (users)
		if (!isset($_SESSION['users'])) $_SESSION['users'] = new cusers();

		// Page ID
		if (!defined("EW_PAGE_ID"))
			define("EW_PAGE_ID", 'add', TRUE);

        // Table name (for backward compatibility)
        if (!defined("EW_TABLE_NAME"))
            define("EW_TABLE_NAME", 'a_customers', TRUE);

        // Start timer
        if (!isset($_SESSION["gTimer"])) $_SESSION["gTimer"] = new cTimer();

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
	var $FormClassName = "form-horizontal ewForm ewAddForm";
	var $DbMasterFilter = "";
	var $DbDetailFilter = "";
	var $StartRec;
	var $Priv = 0;
	var $OldRecordset;
	var $CopyRecord;

	// 
	// Page main
	//
	function Page_Main() {
	    session_start();


		// Process form if post back
		if ($_POST["a_add"] <> "") {
			$this->CurrentAction = $_POST["a_add"]; // Get form action
			$this->CopyRecord = $this->LoadOldRecord(); // Load old recordset
			$this->LoadFormValues(); // Load form values

			// End of modification Permission Access for Export To Feature, by Masino Sinaga, May 5, 2012
		} else { // Not post back

			// Load key values from QueryString
			$this->CopyRecord = TRUE;
			if ($_GET["Customer_ID"] != "") {
				$this->Customer_ID->setQueryStringValue($_GET["Customer_ID"]);
				$this->setKey("Customer_ID", $this->Customer_ID->CurrentValue); // Set up key
			} else {
				$this->setKey("Customer_ID", ""); // Clear key
				$this->CopyRecord = FALSE;
			}
			if ($this->CopyRecord) {
				$this->CurrentAction = "C"; // Copy record
			} else {
				$this->CurrentAction = "I"; // Display blank record
			}
		}

		// Set up Breadcrumb
		$this->SetupBreadcrumb();

		// Set up detail parameters
		$this->SetUpDetailParms();

		// Validate form if post back
		if ($_POST["a_add"] <> "") {
			if (!$this->ValidateForm()) {
				$this->CurrentAction = "I"; // Form error, reset action
				$this->EventCancelled = TRUE; // Event cancelled
				$this->RestoreFormValues(); // Restore form values
				$this->setFailureMessage($_SESSION['gsFormError']);
			}
		} else {
			if ($this->CurrentAction == "I") // Load default values for blank record
				$this->LoadDefaultValues();
		}

		// Perform action based on action code
		switch ($this->CurrentAction) {
			case "I": // Blank record, no action required
				break;
			case "C": // Copy an existing record
				if (!$this->LoadRow()) { // Load record based on key
					if ($this->getFailureMessage() == "") $this->setFailureMessage($_SESSION['Language']->Phrase("NoRecord")); // No record found
					$this->Page_Terminate("a_customerslist.php"); // No matching record, return to list
				}

				// Set up detail parameters
				$this->SetUpDetailParms();
				break;
			case "A": // Add new record
				$this->SendEmail = TRUE; // Send email on add success
				if ($this->AddRow($this->OldRecordset)) { // Add successful

					// Begin of modification Disable Add/Edit Success Message Box, by Masino Sinaga, August 1, 2012
					if (MS_SHOW_ADD_SUCCESS_MESSAGE==TRUE) {
						if ($this->getSuccessMessage() == "")
							$this->setSuccessMessage($_SESSION['Language']->Phrase("AddSuccess")); // Set up success message
					}

					// End of modification Disable Add/Edit Success Message Box, by Masino Sinaga, August 1, 2012
					if ($this->getCurrentDetailTable() <> "") // Master/detail add
						$sReturnUrl = $this->GetDetailUrl();
					else
						$sReturnUrl = $this->getReturnUrl();
					if (ew_GetPageName($sReturnUrl) == "a_customerslist.php")
						$sReturnUrl = $this->AddMasterUrl($this->GetListUrl()); // List page, return to list page with correct master key if necessary
					elseif (ew_GetPageName($sReturnUrl) == "a_customersview.php")
						$sReturnUrl = $this->GetViewUrl(); // View page, return to view page with keyurl directly
					$this->Page_Terminate($sReturnUrl); // Clean up and return
				} else {
					$this->EventCancelled = TRUE; // Event cancelled
					$this->RestoreFormValues(); // Add failed, restore form values

					// Set up detail parameters
					$this->SetUpDetailParms();
				}
				break;
            default:
                echo 'it is not present in any case';
		}

		// Render row based on row type
		$this->RowType = EW_ROWTYPE_ADD; // Render add type

		// Render row
		$this->ResetAttrs();
		$this->RenderRow();
	}

	// Get upload files
	function GetUploadFiles() {

		// Get upload data
	}

	// Load default values
	function LoadDefaultValues() {
		$this->Customer_Number->CurrentValue = NULL;
		$this->Customer_Number->OldValue = $this->Customer_Number->CurrentValue;
		$this->Customer_Name->CurrentValue = NULL;
		$this->Customer_Name->OldValue = $this->Customer_Name->CurrentValue;
		$this->Address->CurrentValue = NULL;
		$this->Address->OldValue = $this->Address->CurrentValue;
		$this->City->CurrentValue = NULL;
		$this->City->OldValue = $this->City->CurrentValue;
		$this->Country->CurrentValue = NULL;
		$this->Country->OldValue = $this->Country->CurrentValue;
		$this->Contact_Person->CurrentValue = NULL;
		$this->Contact_Person->OldValue = $this->Contact_Person->CurrentValue;
		$this->Phone_Number->CurrentValue = NULL;
		$this->Phone_Number->OldValue = $this->Phone_Number->CurrentValue;
		$this->_Email->CurrentValue = NULL;
		$this->_Email->OldValue = $this->_Email->CurrentValue;
		$this->Mobile_Number->CurrentValue = NULL;
		$this->Mobile_Number->OldValue = $this->Mobile_Number->CurrentValue;
		$this->Notes->CurrentValue = NULL;
		$this->Notes->OldValue = $this->Notes->CurrentValue;
		$this->Date_Added->CurrentValue = ew_CurrentDateTime();
		$this->Added_By->CurrentValue = CurrentUserName();
		$this->Date_Updated->CurrentValue = NULL;
		$this->Date_Updated->OldValue = $this->Date_Updated->CurrentValue;
		$this->Updated_By->CurrentValue = NULL;
		$this->Updated_By->OldValue = $this->Updated_By->CurrentValue;
	}

	// Load form values
	function LoadFormValues() {

		// Load from form
		session_start();
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

		$loadoldrecord = $this->LoadOldRecord();
		$loadoldrecord = $loadoldrecord."loadoldrecord";
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
		$this->Date_Added->CurrentValue = $this->Date_Added->FormValue;
		$this->Date_Added->CurrentValue = ew_UnFormatDateTime($this->Date_Added->CurrentValue, 0);
		$this->Added_By->CurrentValue = $this->Added_By->FormValue;
		$this->Date_Updated->CurrentValue = $this->Date_Updated->FormValue;
		$this->Date_Updated->CurrentValue = ew_UnFormatDateTime($this->Date_Updated->CurrentValue, 0);
		$this->Updated_By->CurrentValue = $this->Updated_By->FormValue;
	}






	// Load old record
	function LoadOldRecord() {

		// Load key values from Session
		$bValidKey = TRUE;
		if (strval($this->getKey("Customer_ID")) <> "")
			$this->Customer_ID->CurrentValue = $this->getKey("Customer_ID"); // Customer_ID
		else
			$bValidKey = FALSE;

		// Load old recordset
		if ($bValidKey) {
			$this->CurrentFilter = $this->KeyFilter();
			$sSql = $this->SQL();
			$conn = &$this->Connection();
			$this->OldRecordset = ew_LoadRecordset($sSql, $conn);
			$this->LoadRowValues($this->OldRecordset); // Load row values
		} else {
			$this->OldRecordset = NULL;
		}
		return $bValidKey;
	}

	// Render row values based on field settings
	function RenderRow() {

		// Initialize URLs
		// Call Row_Rendering event

		$this->Row_Rendering();

		// Common render codes for all row types
		// Customer_ID
		// Customer_Number
		// Customer_Name
		// Address
		// City
		// Country
		// Contact_Person
		// Phone_Number
		// Email
		// Mobile_Number
		// Notes
		// Balance
		// Date_Added
		// Added_By
		// Date_Updated
		// Updated_By

		if ($this->RowType == EW_ROWTYPE_VIEW) { // View row

		// Customer_Number
		$this->Customer_Number->ViewValue = $this->Customer_Number->CurrentValue;
		$this->Customer_Number->ViewCustomAttributes = "";

		// Customer_Name
		$this->Customer_Name->ViewValue = $this->Customer_Name->CurrentValue;
		$this->Customer_Name->ViewCustomAttributes = "";

		// Address
		$this->Address->ViewValue = $this->Address->CurrentValue;
		$this->Address->ViewCustomAttributes = "";

		// City
		$this->City->ViewValue = $this->City->CurrentValue;
		$this->City->ViewCustomAttributes = "";

		// Country
		$this->Country->ViewValue = $this->Country->CurrentValue;
		$this->Country->ViewCustomAttributes = "";

		// Contact_Person
		$this->Contact_Person->ViewValue = $this->Contact_Person->CurrentValue;
		$this->Contact_Person->ViewCustomAttributes = "";

		// Phone_Number
		$this->Phone_Number->ViewValue = $this->Phone_Number->CurrentValue;
		$this->Phone_Number->ViewCustomAttributes = "";

		// Email
		$this->_Email->ViewValue = $this->_Email->CurrentValue;
		$this->_Email->ViewCustomAttributes = "";

		// Mobile_Number
		$this->Mobile_Number->ViewValue = $this->Mobile_Number->CurrentValue;
		$this->Mobile_Number->ViewCustomAttributes = "";

		// Notes
		$this->Notes->ViewValue = $this->Notes->CurrentValue;
		$this->Notes->ViewCustomAttributes = "";

		// Balance
		$this->Balance->ViewValue = $this->Balance->CurrentValue;
		$this->Balance->ViewValue = ew_FormatCurrency($this->Balance->ViewValue, 2, -2, -2, -2);
		$this->Balance->CellCssStyle .= "text-align: right;";
		$this->Balance->ViewCustomAttributes = "";

		// Date_Added
		$this->Date_Added->ViewValue = $this->Date_Added->CurrentValue;
		$this->Date_Added->ViewCustomAttributes = "";

		// Added_By
		$this->Added_By->ViewValue = $this->Added_By->CurrentValue;
		$this->Added_By->ViewCustomAttributes = "";

		// Date_Updated
		$this->Date_Updated->ViewValue = $this->Date_Updated->CurrentValue;
		$this->Date_Updated->ViewCustomAttributes = "";

		// Updated_By
		$this->Updated_By->ViewValue = $this->Updated_By->CurrentValue;
		$this->Updated_By->ViewCustomAttributes = "";

			// Customer_Number
			$this->Customer_Number->LinkCustomAttributes = "";
			$this->Customer_Number->HrefValue = "";
			$this->Customer_Number->TooltipValue = "";

			// Customer_Name
			$this->Customer_Name->LinkCustomAttributes = "";
			$this->Customer_Name->HrefValue = "";
			$this->Customer_Name->TooltipValue = "";

			// Address
			$this->Address->LinkCustomAttributes = "";
			$this->Address->HrefValue = "";
			$this->Address->TooltipValue = "";

			// City
			$this->City->LinkCustomAttributes = "";
			$this->City->HrefValue = "";
			$this->City->TooltipValue = "";

			// Country
			$this->Country->LinkCustomAttributes = "";
			$this->Country->HrefValue = "";
			$this->Country->TooltipValue = "";

			// Contact_Person
			$this->Contact_Person->LinkCustomAttributes = "";
			$this->Contact_Person->HrefValue = "";
			$this->Contact_Person->TooltipValue = "";

			// Phone_Number
			$this->Phone_Number->LinkCustomAttributes = "";
			$this->Phone_Number->HrefValue = "";
			$this->Phone_Number->TooltipValue = "";

			// Email
			$this->_Email->LinkCustomAttributes = "";
			$this->_Email->HrefValue = "";
			$this->_Email->TooltipValue = "";

			// Mobile_Number
			$this->Mobile_Number->LinkCustomAttributes = "";
			$this->Mobile_Number->HrefValue = "";
			$this->Mobile_Number->TooltipValue = "";

			// Notes
			$this->Notes->LinkCustomAttributes = "";
			$this->Notes->HrefValue = "";
			$this->Notes->TooltipValue = "";

			// Date_Added
			$this->Date_Added->LinkCustomAttributes = "";
			$this->Date_Added->HrefValue = "";
			$this->Date_Added->TooltipValue = "";

			// Added_By
			$this->Added_By->LinkCustomAttributes = "";
			$this->Added_By->HrefValue = "";
			$this->Added_By->TooltipValue = "";

			// Date_Updated
			$this->Date_Updated->LinkCustomAttributes = "";
			$this->Date_Updated->HrefValue = "";
			$this->Date_Updated->TooltipValue = "";

			// Updated_By
			$this->Updated_By->LinkCustomAttributes = "";
			$this->Updated_By->HrefValue = "";
			$this->Updated_By->TooltipValue = "";
		} elseif ($this->RowType == EW_ROWTYPE_ADD) { // Add row

			// Customer_Number
			$this->Customer_Number->EditAttrs["class"] = "form-control";
			$this->Customer_Number->EditCustomAttributes = "";
			$this->Customer_Number->EditValue = ew_HtmlEncode($this->Customer_Number->CurrentValue);
			$this->Customer_Number->PlaceHolder = ew_RemoveHtml($this->Customer_Number->FldCaption());

			// Customer_Name
			$this->Customer_Name->EditAttrs["class"] = "form-control";
			$this->Customer_Name->EditCustomAttributes = "";
			$this->Customer_Name->EditValue = ew_HtmlEncode($this->Customer_Name->CurrentValue);
			$this->Customer_Name->PlaceHolder = ew_RemoveHtml($this->Customer_Name->FldCaption());

			// Address
			$this->Address->EditAttrs["class"] = "form-control";
			$this->Address->EditCustomAttributes = "";
			$this->Address->EditValue = ew_HtmlEncode($this->Address->CurrentValue);
			$this->Address->PlaceHolder = ew_RemoveHtml($this->Address->FldCaption());

			// City
			$this->City->EditAttrs["class"] = "form-control";
			$this->City->EditCustomAttributes = "";
			$this->City->EditValue = ew_HtmlEncode($this->City->CurrentValue);
			$this->City->PlaceHolder = ew_RemoveHtml($this->City->FldCaption());

			// Country
			$this->Country->EditAttrs["class"] = "form-control";
			$this->Country->EditCustomAttributes = "";
			$this->Country->EditValue = ew_HtmlEncode($this->Country->CurrentValue);
			$this->Country->PlaceHolder = ew_RemoveHtml($this->Country->FldCaption());

			// Contact_Person
			$this->Contact_Person->EditAttrs["class"] = "form-control";
			$this->Contact_Person->EditCustomAttributes = "";
			$this->Contact_Person->EditValue = ew_HtmlEncode($this->Contact_Person->CurrentValue);
			$this->Contact_Person->PlaceHolder = ew_RemoveHtml($this->Contact_Person->FldCaption());

			// Phone_Number
			$this->Phone_Number->EditAttrs["class"] = "form-control";
			$this->Phone_Number->EditCustomAttributes = "";
			$this->Phone_Number->EditValue = ew_HtmlEncode($this->Phone_Number->CurrentValue);
			$this->Phone_Number->PlaceHolder = ew_RemoveHtml($this->Phone_Number->FldCaption());

			// Email
			$this->_Email->EditAttrs["class"] = "form-control";
			$this->_Email->EditCustomAttributes = "";
			$this->_Email->EditValue = ew_HtmlEncode($this->_Email->CurrentValue);
			$this->_Email->PlaceHolder = ew_RemoveHtml($this->_Email->FldCaption());

			// Mobile_Number
			$this->Mobile_Number->EditAttrs["class"] = "form-control";
			$this->Mobile_Number->EditCustomAttributes = "";
			$this->Mobile_Number->EditValue = ew_HtmlEncode($this->Mobile_Number->CurrentValue);
			$this->Mobile_Number->PlaceHolder = ew_RemoveHtml($this->Mobile_Number->FldCaption());

			// Notes
			$this->Notes->EditAttrs["class"] = "form-control";
			$this->Notes->EditCustomAttributes = "";
			$this->Notes->EditValue = ew_HtmlEncode($this->Notes->CurrentValue);
			$this->Notes->PlaceHolder = ew_RemoveHtml($this->Notes->FldCaption());

			// Date_Added
			$this->Date_Added->EditAttrs["class"] = "form-control";
			$this->Date_Added->EditCustomAttributes = "";
			$this->Date_Added->CurrentValue = ew_CurrentDateTime();

			// Added_By
			$this->Added_By->EditAttrs["class"] = "form-control";
			$this->Added_By->EditCustomAttributes = "";
			$this->Added_By->CurrentValue = CurrentUserName();

			// Date_Updated
			// Updated_By
			// Add refer script
			// Customer_Number

			$this->Customer_Number->LinkCustomAttributes = "";
			$this->Customer_Number->HrefValue = "";

			// Customer_Name
			$this->Customer_Name->LinkCustomAttributes = "";
			$this->Customer_Name->HrefValue = "";

			// Address
			$this->Address->LinkCustomAttributes = "";
			$this->Address->HrefValue = "";

			// City
			$this->City->LinkCustomAttributes = "";
			$this->City->HrefValue = "";

			// Country
			$this->Country->LinkCustomAttributes = "";
			$this->Country->HrefValue = "";

			// Contact_Person
			$this->Contact_Person->LinkCustomAttributes = "";
			$this->Contact_Person->HrefValue = "";

			// Phone_Number
			$this->Phone_Number->LinkCustomAttributes = "";
			$this->Phone_Number->HrefValue = "";

			// Email
			$this->_Email->LinkCustomAttributes = "";
			$this->_Email->HrefValue = "";

			// Mobile_Number
			$this->Mobile_Number->LinkCustomAttributes = "";
			$this->Mobile_Number->HrefValue = "";

			// Notes
			$this->Notes->LinkCustomAttributes = "";
			$this->Notes->HrefValue = "";

			// Date_Added
			$this->Date_Added->LinkCustomAttributes = "";
			$this->Date_Added->HrefValue = "";

			// Added_By
			$this->Added_By->LinkCustomAttributes = "";
			$this->Added_By->HrefValue = "";

			// Date_Updated
			$this->Date_Updated->LinkCustomAttributes = "";
			$this->Date_Updated->HrefValue = "";

			// Updated_By
			$this->Updated_By->LinkCustomAttributes = "";
			$this->Updated_By->HrefValue = "";
		}
		if ($this->RowType == EW_ROWTYPE_ADD ||
			$this->RowType == EW_ROWTYPE_EDIT ||
			$this->RowType == EW_ROWTYPE_SEARCH) { // Add / Edit / Search row
			$this->SetupFieldTitles();
		}

		// Call Row Rendered event
		if ($this->RowType <> EW_ROWTYPE_AGGREGATEINIT)
			$this->Row_Rendered();
	}

	// Add record
	function AddRow($rsold = NULL) {
	    session_start();

		$conn = &$this->Connection();

		// Begin transaction
		if ($this->getCurrentDetailTable() <> "")
			$conn->BeginTrans();

		// Load db values from rsold
		if ($rsold) {
			$this->LoadDbValues($rsold);
		}
		$rsnew = array();

		// Customer_Number
		$this->Customer_Number->SetDbValueDef($rsnew, $this->Customer_Number->CurrentValue, "", FALSE);

		// Customer_Name
		$this->Customer_Name->SetDbValueDef($rsnew, $this->Customer_Name->CurrentValue, "", FALSE);

		// Address
		$this->Address->SetDbValueDef($rsnew, $this->Address->CurrentValue, "", FALSE);

		// City
		$this->City->SetDbValueDef($rsnew, $this->City->CurrentValue, "", FALSE);

		// Country
		$this->Country->SetDbValueDef($rsnew, $this->Country->CurrentValue, "", FALSE);

		// Contact_Person
		$this->Contact_Person->SetDbValueDef($rsnew, $this->Contact_Person->CurrentValue, "", FALSE);

		// Phone_Number
		$this->Phone_Number->SetDbValueDef($rsnew, $this->Phone_Number->CurrentValue, "", FALSE);

		// Email
		$this->_Email->SetDbValueDef($rsnew, $this->_Email->CurrentValue, "", FALSE);

		// Mobile_Number
		$this->Mobile_Number->SetDbValueDef($rsnew, $this->Mobile_Number->CurrentValue, "", FALSE);

		// Notes
		$this->Notes->SetDbValueDef($rsnew, $this->Notes->CurrentValue, "", FALSE);

		// Date_Added
		$this->Date_Added->SetDbValueDef($rsnew, $this->Date_Added->CurrentValue, NULL, FALSE);

		// Added_By
		$this->Added_By->SetDbValueDef($rsnew, $this->Added_By->CurrentValue, NULL, FALSE);

		// Date_Updated
		$this->Date_Updated->SetDbValueDef($rsnew, ew_CurrentDateTime(), NULL);
		$rsnew['Date_Updated'] = &$this->Date_Updated->DbValue;

		// Updated_By
		$this->Updated_By->SetDbValueDef($rsnew, CurrentUserName(), NULL);
		$rsnew['Updated_By'] = &$this->Updated_By->DbValue;

		// Call Row Inserting event
		$rs = ($rsold == NULL) ? NULL : $rsold->fields;
		$bInsertRow = $this->Row_Inserting($rs);
		if ($bInsertRow) {
			$conn->raiseErrorFn = $_SESSION['EW_ERROR_FN']; // v11.0.4
			$AddRow = $this->Insert($rsnew);
			$conn->raiseErrorFn = '';
			if ($AddRow) {

				// Get insert id if necessary
				$this->Customer_ID->setDbValue($conn->Insert_ID());
				$rsnew['Customer_ID'] = $this->Customer_ID->DbValue;
			}
		} else {
			if ($this->getSuccessMessage() <> "" || $this->getFailureMessage() <> "") {

				// Use the message, do nothing
			} elseif ($this->CancelMessage <> "") {
				$this->setFailureMessage($this->CancelMessage);
				$this->CancelMessage = "";
			} else {
				$this->setFailureMessage($_SESSION['Language']->Phrase("InsertCancelled"));
			}
			$AddRow = FALSE;
		}



		// Add detail records
		if ($AddRow) {
			$DetailTblVar = explode(",", $this->getCurrentDetailTable());
			if (in_array('a_sales', $DetailTblVar) && $_SESSION['a_sales']->DetailAdd) {
				$_SESSION['a_sales']->Customer_ID->setSessionValue($this->Customer_Number->CurrentValue); // Set master key
				if (!isset($_SESSION['a_sales_grid'])) $_SESSION['a_sales_grid'] = new ca_sales_grid(); // Get detail page object
				$AddRow = $_SESSION['a_sales_grid']->GridInsert();
				if (!$AddRow)
					$_SESSION['a_sales']->Customer_ID->setSessionValue(""); // Clear master key if insert failed
			}
		}

		// Commit/Rollback transaction
		if ($this->getCurrentDetailTable() <> "") {
			if ($AddRow) {
				$conn->CommitTrans(); // Commit transaction
			} else {
				$conn->RollbackTrans(); // Rollback transaction
			}
		}
		if ($AddRow) {

			// Call Row Inserted event
			$rs = ($rsold == NULL) ? NULL : $rsold->fields;
			$this->Row_Inserted();
		}
		return $AddRow;
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
				if ($_SESSION['a_sales_grid']->DetailAdd) {
					if ($this->CopyRecord)
						$_SESSION['a_sales_grid']->CurrentMode = "copy";
					else
						$_SESSION['a_sales_grid']->CurrentMode = "add";
					$_SESSION['a_sales_grid']->CurrentAction = "gridadd";

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

		$Breadcrumb = new cBreadcrumb();
		$url = substr(ew_CurrentUrl(), strrpos(ew_CurrentUrl(), "/")+1); // v11.0.4
		$Breadcrumb->Add("list", $this->TableVar, $this->AddMasterUrl("a_customerslist.php"), "", $this->TableVar, TRUE);
		$PageId = ($this->CurrentAction == "C") ? "Copy" : "Add";
		$Breadcrumb->Add("add", $PageId, $url); // v11.0.4
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
			//$msg = "your failure message";
		} elseif ($type == 'warning') {
            $msg = $msg."";
			//$msg = "your warning message";
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
if (!isset($a_customers_add)) $a_customers_add = new ca_customers_add();

// Page init
$a_customers_add->Page_Init();

// Page main
$a_customers_add->Page_Main();

// Begin of modification Displaying Breadcrumb Links in All Pages, by Masino Sinaga, May 4, 2012
getCurrentPageTitle(ew_CurrentPage());

// End of modification Displaying Breadcrumb Links in All Pages, by Masino Sinaga, May 4, 2012
// Global Page Rendering event (in userfn*.php)

Page_Rendering();

// Global auto switch table width style (in userfn*.php), by Masino Sinaga, January 7, 2015
AutoSwitchTableWidthStyle();

// Page Rendering event
$a_customers_add->Page_Render();
?>
<?php include_once "header.php" ?>
<script type="text/javascript">

// Form object
var CurrentPageID = EW_PAGE_ID = "add";
var CurrentForm = fa_customersadd = new ew_Form("fa_customersadd", "add");



// Form_CustomValidate event
fa_customersadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!

 	// Your custom validation code here, return false if invalid. 
 	return true;
 }

// Use JavaScript validation or not
<?php if (EW_CLIENT_VALIDATE) { ?>
fa_customersadd.ValidateRequired = true;
<?php } else { ?>
fa_customersadd.ValidateRequired = false; 
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
<?php $a_customers_add->ShowPageHeader(); ?>
<?php
$a_customers_add->ShowMessage();
?>
<form name="fa_customersadd" id="fa_customersadd" class="<?php echo $a_customers_add->FormClassName ?>" action="<?php echo ew_CurrentPage() ?>" method="post">
<?php if ($a_customers_add->CheckToken) { ?>
<input type="hidden" name="<?php echo EW_TOKEN_NAME ?>" value="<?php echo $a_customers_add->Token ?>">
<?php } ?>
<input type="hidden" name="t" value="a_customers">
<input type="hidden" name="a_add" id="a_add" value="A">
<div>

    <?php
        $html_if = new html_if();
        $html_if->html_code($a_customers,$Language);
    ?>

<span id="el_a_customers_Date_Added">
<input type="hidden" data-table="a_customers" data-field="x_Date_Added" name="x_Date_Added" id="x_Date_Added" value="<?php echo ew_HtmlEncode($a_customers->Date_Added->CurrentValue) ?>">
</span>
<span id="el_a_customers_Added_By">
<input type="hidden" data-table="a_customers" data-field="x_Added_By" name="x_Added_By" id="x_Added_By" value="<?php echo ew_HtmlEncode($a_customers->Added_By->CurrentValue) ?>">
</span>
</div>
<?php
	if (in_array($_SESSION['a_sales'], explode(",", $a_customers->getCurrentDetailTable())) && $a_sales->DetailAdd) {
?>
<?php if ($a_customers->getCurrentDetailTable() <> "") { ?>
<h4 class="ewDetailCaption"><?php echo $Language->TablePhrase("a_sales", "TblCaption") ?></h4>
<?php } ?>
<?php include_once "a_salesgrid.php" ?>
<?php } ?>
<div class="form-group">
	<div class="col-sm-offset-4 col-sm-8">
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><?php echo $Language->Phrase("AddBtn") ?></button>
<button class="btn btn-danger ewButton" name="btnCancel" id="btnCancel" type="button" data-href="<?php echo $a_customers_add->getReturnUrl() ?>"><?php echo $Language->Phrase("CancelBtn") ?></button>
	</div>
</div>
</form>
<script type="text/javascript">
fa_customersadd.Init();
</script>
<?php
$a_customers_add->ShowPageFooter();
if (EW_DEBUG_ENABLED)
	echo ew_DebugMsg();
?>
<script type="text/javascript">

// Write your table-specific startup script here
// document.write("page loaded");

</script>
<?php if (MS_ENTER_MOVING_CURSOR_TO_NEXT_FIELD) { ?>
<script type="text/javascript">
$(document).ready(function(){$("#fa_customersadd:first *:input[type!=hidden]:first").focus(),$("input").keydown(function(i){if(13==i.which){var e=$(this).closest("form").find(":input:visible:enabled"),n=e.index(this);n==e.length-1||(e.eq(e.index(this)+1).focus(),i.preventDefault())}else 113==i.which&&$("#btnAction").click()}),$("select").keydown(function(i){if(13==i.which){var e=$(this).closest("form").find(":input:visible:enabled"),n=e.index(this);n==e.length-1||(e.eq(e.index(this)+1).focus(),i.preventDefault())}else 113==i.which&&$("#btnAction").click()}),$("radio").keydown(function(i){if(13==i.which){var e=$(this).closest("form").find(":input:visible:enabled"),n=e.index(this);n==e.length-1||(e.eq(e.index(this)+1).focus(),i.preventDefault())}else 113==i.which&&$("#btnAction").click()})});
</script>
<?php } ?>
<?php if ($a_customers->Export == "") { ?>
<script type="text/javascript">
$('#btnAction').attr('onclick', 'return alertifyAdd(this)'); function alertifyAdd(obj) { <?php global $Language; ?> if (fa_customersadd.Validate() == true ) { alertify.confirm("<?php echo $Language->Phrase('AlertifyAddConfirm'); ?>", function (e) { if (e) { $(window).unbind('beforeunload'); alertify.success("<?php echo $Language->Phrase('AlertifyAdd'); ?>"); $("#fa_customersadd").submit(); } }).set("title", "<?php echo $Language->Phrase('AlertifyConfirm'); ?>").set("defaultFocus", "cancel").set('oncancel', function(closeEvent){ alertify.error('<?php echo $Language->Phrase('AlertifyCancel'); ?>');}).set('labels', {ok:'<?php echo $Language->Phrase("MyOKMessage"); ?>!', cancel:'<?php echo $Language->Phrase("MyCancelMessage"); ?>'}); } return false; }
</script>
<?php } ?>
<?php include_once "footer.php" ?>
<?php
$a_customers_add->Page_Terminate();
?>
