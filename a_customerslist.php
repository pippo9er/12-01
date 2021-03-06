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
<?php include_once "restoreFilterList.php" ?>
<?php include_once "setup_option.php" ?>
<?php include_once "render_list.php" ?>
<?php include_once "search2.php"?>
<?php include_once "function_1.php"?>
<?php include_once "page_main.php"?>
<?php include_once "build_basic_search.php"?>
<?php include_once "list_action.php"?>
<?php include_once "export_email.php"?>
<?php include_once "page_init.php"?>
<?php include_once "cod_html.php"?>
<?php include_once "render_option_ext.php"?>
<?php include_once "funct_2.php"?>
<?php include_once "cod_html_5.php"?>
<?php include_once "cod_html_6.php"?>
<?php

//
// Page class
//

$a_customers_list = NULL; // Initialize page object first

class ca_customers_list extends ca_customers {

	// Page ID
	var $PageID = 'list';

	// Project ID
	var $ProjectID = "{B36B93AF-B58F-461B-B767-5F08C12493E9}";

	// Table name
	var $TableName = 'a_customers';

	// Page object name
	var $PageObjName = 'a_customers_list';

	// Grid form hidden field names
	var $FormName = 'fa_customerslist';
	var $FormActionName = 'k_action';
	var $FormKeyName = 'k_key';
	var $FormOldKeyName = 'k_oldkey';
	var $FormBlankRowName = 'k_blankrow';
	var $FormKeyCountName = 'key_count';

	// Page name

	// Page URLs
	var $AddUrl;
	var $EditUrl;
	var $CopyUrl;
	var $DeleteUrl;
	var $ViewUrl;
	var $ListUrl;

	// Export URLs
	var $ExportPrintUrl;
	var $ExportHtmlUrl;
	var $ExportExcelUrl;
	var $ExportWordUrl;
	var $ExportXmlUrl;
	var $ExportCsvUrl;
	var $ExportPdfUrl;

	// Custom export
	var $ExportExcelCustom = FALSE;
	var $ExportWordCustom = FALSE;
	var $ExportPdfCustom = FALSE;
	var $ExportEmailCustom = FALSE;

	// Update URLs
	var $InlineAddUrl;
	var $InlineCopyUrl;
	var $InlineEditUrl;
	var $GridAddUrl;
	var $GridEditUrl;
	var $MultiDeleteUrl;
	var $MultiUpdateUrl;


	var $PageHeader;
	var $PageFooter;

	// Show Page Header
	function ShowPageHeader() {
		$sHeader = $this->PageHeader;
		$this->Page_DataRendering();
		if ($sHeader <> "") { // Header exists, display
            $stg="<p>" . $sHeader . "</p>";
            echo $stg;
		}
	}

	// Show Page Footer
	function ShowPageFooter() {
		$sFooter = $this->PageFooter;
		$this->Page_DataRendered();
		if ($sFooter <> "") { // Footer exists, display
            $stg="<p>" . $sFooter . "</p>";
            echo $stg;
		}
	}

	// Validate page request
	function IsPageRequest() {
	    session_start();
		if ($this->UseTokenInUrl) {
			if ($_SESSION['objForm'])
				return ($this->TableVar == $_SESSION['objForm']->GetValue("t"));
			if ($_GET["t"] <> "")
				return ($this->TableVar == $_GET["t"]);
		} else {
			return TRUE;
		}
	}
	var $Token = "";
	var $TokenTimeout = 0;
	var $CheckToken = EW_CHECK_TOKEN;
	var $CheckTokenFn = "ew_CheckToken";
	var $CreateTokenFn = "ew_CreateToken";

	// Valid Post
	function ValidPost() {
		if (!$this->CheckToken || !ew_IsHttpPost())
			return TRUE;
		if (!isset($_POST[EW_TOKEN_NAME]))
			return FALSE;
		$fn = $this->CheckTokenFn;
		if (is_callable($fn))
			return $fn($_POST[EW_TOKEN_NAME], $this->TokenTimeout);
		return FALSE;
	}

	// Create Token
	function CreateToken() {
		session_start();
		if ($this->CheckToken) {
			$fn = $this->CreateTokenFn;
			if ($this->Token == "" && is_callable($fn)) // Create token
				$this->Token = $fn();
            $_SESSION['gsToken'] = $this->Token; 
		}
	}

	//
	// Page class constructor
	//
	function __construct() {

        session_start();


		$_SESSION["Page"] = &$this;
		$this->TokenTimeout = ew_SessionTimeoutTime();

		// Language object
        $_SESSION['Language']  = new cLanguage();

		// Parent constuctor
		parent::__construct();

		// Table object (a_customers)
		if (!isset($_SESSION["a_customers"]) || get_class($_SESSION["a_customers"]) == "ca_customers") {
            $_SESSION["a_customers"] = &$this;
            $_SESSION["Table"] = &$GLOBALS["a_customers"];
		}

		// Initialize URLs
		$this->ExportPrintUrl = $this->PageUrl() . "export=print";
		$this->ExportExcelUrl = $this->PageUrl() . "export=excel";
		$this->ExportWordUrl = $this->PageUrl() . "export=word";
		$this->ExportHtmlUrl = $this->PageUrl() . "export=html";
		$this->ExportXmlUrl = $this->PageUrl() . "export=xml";
		$this->ExportCsvUrl = $this->PageUrl() . "export=csv";
		$this->ExportPdfUrl = $this->PageUrl() . "export=pdf";
		$this->AddUrl = "a_customersadd.php?" . EW_TABLE_SHOW_DETAIL . "=";
		$this->InlineAddUrl = $this->PageUrl() . "a=add";
		$this->GridAddUrl = $this->PageUrl() . "a=gridadd";
		$this->GridEditUrl = $this->PageUrl() . "a=gridedit";
		$this->MultiDeleteUrl = "a_customersdelete.php";
		$this->MultiUpdateUrl = "a_customersupdate.php";

		// Table object (users)
		if (!isset($_SESSION['users'])) $_SESSION['users'] = new cusers();

		// Page ID
		if (!defined("EW_PAGE_ID"))
			define("EW_PAGE_ID", 'list', TRUE);

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

		//}

		// List options
		$this->ListOptions = new cListOptions();
		$this->ListOptions->TableVar = $this->TableVar;

		// Export options
		$this->ExportOptions = new cListOptions();
		$this->ExportOptions->Tag = "div";
		$this->ExportOptions->TagClassName = "ewExportOption";

		// Other options
		$this->OtherOptions['addedit'] = new cListOptions();
		$this->OtherOptions['addedit']->Tag = "div";
		$this->OtherOptions['addedit']->TagClassName = "ewAddEditOption";
		$this->OtherOptions['detail'] = new cListOptions();
		$this->OtherOptions['detail']->Tag = "div";
		$this->OtherOptions['detail']->TagClassName = "ewDetailOption";
		$this->OtherOptions['action'] = new cListOptions();
		$this->OtherOptions['action']->Tag = "div";
		$this->OtherOptions['action']->TagClassName = "ewActionOption";

		// Filter options
		$this->FilterOptions = new cListOptions();
		$this->FilterOptions->Tag = "div";
		$this->FilterOptions->TagClassName = "ewFilterOption fa_customerslistsrch";

		// List actions
		$this->ListActions = new cListActions();

	}

	// 
	//  Page_Init
	//
	function Page_Init() {
	
        session_start();
        //$_SESSION['UserTableConn']= $UserTableConn;

		if (!isset($_SESSION['table_a_customers_views'])) { 
			$_SESSION['table_a_customers_views'] = 0;
		}
		$_SESSION['table_a_customers_views'] = $_SESSION['table_a_customers_views']+1;

		// User profile
		$_SESSION['UserProfile'] = new cUserProfile();

		// Security
        $_SESSION['Security'] = new cAdvancedSecurity();
		if (IsPasswordExpired())
			$this->Page_Terminate(ew_GetUrl("changepwd.php"));
		if (!$_SESSION['Security']->IsLoggedIn()) $_SESSION['Security']->AutoLogin();
		if ($_SESSION['Security']->IsLoggedIn()) $_SESSION['Security']->TablePermission_Loading();
		$_SESSION['Security']->LoadCurrentUserLevel($this->ProjectID . $this->TableName);
		if ($_SESSION['Security']->IsLoggedIn()) $_SESSION['Security']->TablePermission_Loaded();
		if (!$_SESSION['Security']->CanList()) {
			$_SESSION['Security']->SaveLastUrl();
			$this->setFailureMessage($_SESSION['Language']->Phrase("NoPermission")); // Set no permission
			$this->Page_Terminate(ew_GetUrl("index.php"));
		}

		// Begin of modification Auto Logout After Idle for the Certain Time, by Masino Sinaga, May 5, 2012
		if (IsLoggedIn() && !IsSysAdmin()) {

			// Begin of modification by Masino Sinaga, May 25, 2012 in order to not autologout after clear another user's session ID whenever back to another page.           
			$_SESSION['UserProfile']->LoadProfileFromDatabase(CurrentUserName());

			// End of modification by Masino Sinaga, May 25, 2012 in order to not autologout after clear another user's session ID whenever back to another page.
			// Begin of modification Save Last Users' Visitted Page, by Masino Sinaga, May 25, 2012

			$lastpage = ew_CurrentPage();
			if ($lastpage!='logout.php' && $lastpage!='index.php') {
				$lasturl = ew_CurrentUrl();
				$sFilterUserID = str_replace("%u", ew_AdjustSql(CurrentUserName(), EW_USER_TABLE_DBID), EW_USER_NAME_FILTER);
				ew_Execute("UPDATE ".EW_USER_TABLE." SET Current_URL = '".$lasturl."' WHERE ".$sFilterUserID."", $_SESSION['$UserTableConn']);
			}

			// End of modification Save Last Users' Visitted Page, by Masino Sinaga, May 25, 2012
			$LastAccessDateTime = strval($_SESSION['UserProfile']->Profile[EW_USER_PROFILE_LAST_ACCESSED_DATE_TIME]);
			$nDiff = intval(ew_DateDiff($LastAccessDateTime, ew_StdCurrentDateTime(), "s"));
			$nCons = intval(MS_AUTO_LOGOUT_AFTER_IDLE_IN_MINUTES) * 60;
			if ($nDiff > $nCons) {

				//header("Location: logout.php?expired=1");
			}
		}

		// End of modification Auto Logout After Idle for the Certain Time, by Masino Sinaga, May 5, 2012
		// Update last accessed time

		if ($_SESSION['UserProfile']->IsValidUser(CurrentUserName(), session_id())) {

			// Do nothing since it's a valid user! SaveProfileToDatabase has been handled from IsValidUser method of UserProfile object.
		} else {

			// Begin of modification How to Overcome "User X already logged in" Issue, by Masino Sinaga, July 22, 2014
			// echo $_SESSION['Language']->Phrase("UserProfileCorrupted");

			header("Location: logout.php");

			// End of modification How to Overcome "User X already logged in" Issue, by Masino Sinaga, July 22, 2014
		}
		if (MS_USE_CONSTANTS_IN_CONFIG_FILE == FALSE) {

			// Call this new function from userfn*.php file
			My_Global_Check();
		}

		// Get export parameters
		$custom = "";
		if ($_GET["export"] <> "") {
			$this->Export = $_GET["export"];
			$custom = $_GET["custom"];
		} elseif ($_POST["export"] <> "") {
			$this->Export = $_POST["export"];
			$custom = $_POST["custom"];
		} elseif (ew_IsHttpPost()) {
			if ($_POST["exporttype"] <> "")
				$this->Export = $_POST["exporttype"];
			$custom = $_POST["custom"];
		} else {
			$this->setExportReturnUrl(ew_CurrentUrl());
		}
		$_SESSION['gsExportFile'] = $this->TableVar; // Get export file, used in header

		// Begin of modification Permission Access for Export To Feature, by Masino Sinaga, To prevent users entering from URL, May 12, 2012

		if ($_SESSION['gsExport']=="print") {
			if (!$_SESSION['Security']->CanExportToPrint() && !$_SESSION['Security']->IsAdmin()) {
				echo $_SESSION['Language']->Phrase("nopermission");
				print "Error: No permission!";
			}
		} elseif ($_SESSION['gsExport']=="excel") {
			if (!$_SESSION['Security']->CanExportToExcel() && !$_SESSION['Security']->IsAdmin()) {
				echo $_SESSION['Language']->Phrase("nopermission");
                print "Error: No permission!";
			}   
		} elseif ($_SESSION['gsExport']=="word") {
			if (!$_SESSION['Security']->CanExportToWord() && !$_SESSION['Security']->IsAdmin()) {
				echo $_SESSION['Language']->Phrase("nopermission");
                print "Error: No permission!";
			}   
		} elseif ($_SESSION['gsExport']=="html") {
			if (!$_SESSION['Security']->CanExportToHTML() && !$_SESSION['Security']->IsAdmin()) {
				echo $_SESSION['Language']->Phrase("nopermission");
                print "Error: No permission!";
			}   
		} elseif ($_SESSION['gsExport']=="csv") {
			if (!$_SESSION['Security']->CanExportToCSV() && !$_SESSION['Security']->IsAdmin()) {
				echo $_SESSION['Language']->Phrase("nopermission");
                print "Error: No permission!";
			}   
		} elseif ($_SESSION['gsExport']=="xml") {
			if (!$_SESSION['Security']->CanExportToXML() && !$_SESSION['Security']->IsAdmin()) {
				echo $_SESSION['Language']->Phrase("nopermission");
                print "Error: No permission!";
			}   
		} elseif ($_SESSION['gsExport']=="pdf") {
			if (!$_SESSION['Security']->CanExportToPDF() && !$_SESSION['Security']->IsAdmin()) {
				echo $_SESSION['Language']->Phrase("nopermission");
                print "Error: No permission!";
			}   
		} elseif ($_SESSION['gsExport']=="email") {
			if (!$_SESSION['Security']->CanExportToEmail() && !$_SESSION['Security']->IsAdmin()) {
				echo $_SESSION['Language']->Phrase("nopermission");
                print "Error: No permission!";
			}   
		}

		// End of modification Permission Access for Export To Feature, by Masino Sinaga, To prevent users entering from URL, May 12, 2012
		// Get custom export parameters

		if ($this->Export <> "" && $custom <> "") {
			$this->CustomExport = $this->Export;
			$this->Export = "print";
		}
		$_SESSION['gsCustomExport'] = $this->CustomExport;
		$_SESSION['gsExport'] = $this->Export; // Get export parameter, used in header

		// Update Export URLs
		if (defined("EW_USE_PHPEXCEL"))
			$this->ExportExcelCustom = FALSE;
		if ($this->ExportExcelCustom)
			$this->ExportExcelUrl .= "&amp;custom=1";
		if (defined("EW_USE_PHPWORD"))
			$this->ExportWordCustom = FALSE;
		if ($this->ExportWordCustom)
			$this->ExportWordUrl .= "&amp;custom=1";
		if ($this->ExportPdfCustom)
			$this->ExportPdfUrl .= "&amp;custom=1";
		$this->CurrentAction = ($_GET["a"] <> "") ? $_GET["a"] : $_POST["a_list"]; // Set up current action

		// Get grid add count
		$gridaddcnt = $_GET[EW_TABLE_GRID_ADD_ROW_COUNT];
		if (is_numeric($gridaddcnt) && $gridaddcnt > 0)
			$this->GridAddRowCount = $gridaddcnt;

		// Set up list options
		$this->SetupListOptions();

		// Setup export options
		$this->SetupExportOptions();

		// Global Page Loading event (in userfn*.php)
		Page_Loading();

// Begin of modification Disable/Enable Registration Page, by Masino Sinaga, May 14, 2012
// End of modification Disable/Enable Registration Page, by Masino Sinaga, May 14, 2012
		// Page Load event

		$this->Page_Load();

		// Check token
		if (!$this->ValidPost()) {
			echo $_SESSION['Language']->Phrase("InvalidPostRequest");
			$this->Page_Terminate();
			print "Error: invalid post request";
		}
		if (ALWAYS_COMPARE_ROOT_URL == TRUE) {
			if ($_SESSION['php_stock_Root_URL'] <> Get_Root_URL()) {
                $strDest = check(rawurlencode($_SESSION['php_stock_Root_URL']));
				header("Location: " . $strDest);
			}
		}

		// Process auto fill
		if ($_POST["ajax"] == "autofill") {

			// Process auto fill for detail table 'a_sales'
			if ($_POST["grid"] == "fa_salesgrid") {
				if (!isset($GLOBALS["a_sales_grid"])) $GLOBALS["a_sales_grid"] = new ca_sales_grid;
				$GLOBALS["a_sales_grid"]->Page_Init();
				$this->Page_Terminate();
				print "Error: failed process for detail table 'a_sales'";
			}
			$results = $this->GetAutoFill();
			if ($results) {

				// Clean output buffer
				if (!EW_DEBUG_ENABLED && ob_get_length())
					ob_end_clean();
				echo $results;
				$this->Page_Terminate();
				print "Error: clean output buffer";
			}
		}

		// Create Token
		$this->CreateToken();

		// Setup other options
		$this->SetupOtherOptions();

		// Set up custom action (compatible with old version)
		foreach ($this->CustomActions as $name => $action)
			$this->ListActions->Add($name, $action);

		// Show checkbox column if multiple action
		foreach ($this->ListActions->Items as $listaction) {
			if ($listaction->Select == EW_ACTION_MULTIPLE && $listaction->Allow) {
				$this->ListOptions->Items["checkbox"]->Visible = TRUE;
				break;
			}
		}

		$init= new page_init();
		$init->Page_Init1($this);

	}

	//
	// Page_Terminate
	//
	function Page_Terminate($url = "") {

		// Page Unload event
		$this->Page_Unload();

		// Global Page Unloaded event (in userfn*.php)
		Page_Unloaded();

		// Export
		
		if ($this->CustomExport <> "" && $this->CustomExport == $this->Export && array_key_exists($this->CustomExport, $_SESSION['EW_EXPORT'])) {
				$sContent = ob_get_contents();
			if ($_SESSION['gsExportFile'] == "") $_SESSION['gsExportFile'] = $this->TableVar;
			$class = $_SESSION['EW_EXPORT'][$this->CustomExport];
			if (class_exists($class)) {
				$doc = new $class($_SESSION['$a_customers']);
				$doc->Text = $sContent;
				if ($this->Export == "email")
					echo $this->ExportEmail($doc->Text);
				else
					$doc->Export();
				ew_DeleteTmpImages(); // Delete temp images
				print "Error: failed export ";
			}
		}
		$this->Page_Redirecting();

		 // Close connection
		ew_CloseConn();

		// Go to URL if specified
		if ($url <> "") {
			if (!EW_DEBUG_ENABLED && ob_get_length())
				ob_end_clean();
			header("Location: " . $url);
		}
		print "Error: invalid URL specified";
	}

	// Class variables
	var $ListOptions; // List options
	var $ExportOptions; // Export options
	var $SearchOptions; // Search options
	var $OtherOptions = array(); // Other options
	var $FilterOptions; // Filter options
	var $ListActions; // List actions
	var $SelectedCount = 0;
	var $SelectedIndex = 0;

// Begin of modification Customize Navigation/Pager Panel, by Masino Sinaga, May 2, 2012
    var $DisplayRecs = MS_TABLE_RECPERPAGE_VALUE;

// End of modification Customize Navigation/Pager Panel, by Masino Sinaga, May 2, 2012
	var $SearchPanelCollapsed = TRUE; // Modified by Masino Sinaga, September 23, 2014
	var $StartRec;
	var $StopRec;
	var $TotalRecs = 0;
	var $RecRange = 10;
	var $Pager;
	var $DefaultSearchWhere = ""; // Default search WHERE clause
	var $SearchWhere = ""; // Search WHERE clause
	var $RecCnt = 0; // Record count
	var $EditRowCnt;
	var $StartRowCnt = 1;
	var $RowCnt = 0;
	var $Attrs = array(); // Row attributes and cell attributes
	var $RowIndex = 0; // Row index
	var $KeyCount = 0; // Key count
	var $RowAction = ""; // Row action
	var $RowOldKey = ""; // Row old key (for copy)
	var $RecPerRow = 0;
	var $MultiColumnClass;
	var $MultiColumnEditClass = "col-sm-12";
	var $MultiColumnCnt = 12;
	var $MultiColumnEditCnt = 12;
	var $GridCnt = 0;
	var $ColCnt = 0;
	var $DbMasterFilter = ""; // Master filter
	var $DbDetailFilter = ""; // Detail filter
	var $MasterRecordExists;	
	var $MultiSelectKey;
	var $Command;
	var $RestoreSearch = FALSE;
	var $a_sales_Count;
	var $DetailPages;
	var $Recordset;
	var $OldRecordset;

	//
	// Page main
	//
	function Page_Main() {
		$pmain= new page_main();
		$pmain->PageMain1($this);
	}

	// Set up number of records displayed per page
	function SetUpDisplayRecs() {

	// Begin of modification Customize Navigation/Pager Panel, by Masino Sinaga, May 2, 2012

        $sWrk = $_GET[EW_TABLE_REC_PER_PAGE];
        if ($sWrk > MS_TABLE_MAXIMUM_SELECTED_RECORDS || strtolower($sWrk) == "all") {
            $sWrk = MS_TABLE_MAXIMUM_SELECTED_RECORDS;
            $this->setFailureMessage(str_replace("%t", MS_TABLE_MAXIMUM_SELECTED_RECORDS, $_SESSION['Language']->Phrase("MaximumRecordsPerPage")));
        }

	// End of modification Customize Navigation/Pager Panel, by Masino Sinaga, May 2, 2012
		if ($sWrk <> "") {
			if (is_numeric($sWrk)) {
				$this->DisplayRecs = intval($sWrk);
			} else {
				if (strtolower($sWrk) == "all") { // Display all records
					$this->DisplayRecs = -1;
				} else {
					$this->DisplayRecs = 20; // Non-numeric, load default
				}
			}
			$this->setRecordsPerPage($this->DisplayRecs); // Save to Session

			// Reset start position
			$this->StartRec = 1;
			$this->setStartRecordNumber($this->StartRec);
		}
	}

	// Build filter for all keys
	function BuildKeyFilter() {
		
		$sWrkFilter = "";

		// Update row index and get row key
		$rowindex = 1;
		$_SESSION['$objForm']->Index = $rowindex;
		$sThisKey = strval($_SESSION['$objForm']->GetValue($this->FormKeyName));
		while ($sThisKey <> "") {
			if ($this->SetupKeyValues($sThisKey)) {
				$sFilter = $this->KeyFilter();
				if ($sWrkFilter <> "") $sWrkFilter .= " OR ";
				$sWrkFilter .= $sFilter;
			} else {
				$sWrkFilter = "0=1";
				break;
			}

			// Update row index and get row key
			$rowindex++; // Next row
			$_SESSION['$objForm']->Index = $rowindex;
			$sThisKey = strval($_SESSION['$objForm']->GetValue($this->FormKeyName));
		}
		return $sWrkFilter;
	}

	// Set up key values
	function SetupKeyValues($key) {
		$arrKeyFlds = explode($GLOBALS["EW_COMPOSITE_KEY_SEPARATOR"], $key);
		if (count($arrKeyFlds) >= 1) {
			$this->Customer_ID->setFormValue($arrKeyFlds[0]);
			if (!is_numeric($this->Customer_ID->FormValue))
				return FALSE;
		}
		return TRUE;
	}

	// Get list of filters
	function GetFilterList() {

		// Initialize
		$sFilterList = "";
		$sFilterList = ew_Concat($sFilterList, $this->Customer_ID->AdvancedSearch->ToJSON(), ","); // Field Customer_ID
		$sFilterList = ew_Concat($sFilterList, $this->Customer_Number->AdvancedSearch->ToJSON(), ","); // Field Customer_Number
		$sFilterList = ew_Concat($sFilterList, $this->Customer_Name->AdvancedSearch->ToJSON(), ","); // Field Customer_Name
		$sFilterList = ew_Concat($sFilterList, $this->Address->AdvancedSearch->ToJSON(), ","); // Field Address
		$sFilterList = ew_Concat($sFilterList, $this->City->AdvancedSearch->ToJSON(), ","); // Field City
		$sFilterList = ew_Concat($sFilterList, $this->Country->AdvancedSearch->ToJSON(), ","); // Field Country
		$sFilterList = ew_Concat($sFilterList, $this->Contact_Person->AdvancedSearch->ToJSON(), ","); // Field Contact_Person
		$sFilterList = ew_Concat($sFilterList, $this->Phone_Number->AdvancedSearch->ToJSON(), ","); // Field Phone_Number
		$sFilterList = ew_Concat($sFilterList, $this->_Email->AdvancedSearch->ToJSON(), ","); // Field Email
		$sFilterList = ew_Concat($sFilterList, $this->Mobile_Number->AdvancedSearch->ToJSON(), ","); // Field Mobile_Number
		$sFilterList = ew_Concat($sFilterList, $this->Notes->AdvancedSearch->ToJSON(), ","); // Field Notes
		$sFilterList = ew_Concat($sFilterList, $this->Balance->AdvancedSearch->ToJSON(), ","); // Field Balance
		$sFilterList = ew_Concat($sFilterList, $this->Date_Added->AdvancedSearch->ToJSON(), ","); // Field Date_Added
		$sFilterList = ew_Concat($sFilterList, $this->Added_By->AdvancedSearch->ToJSON(), ","); // Field Added_By
		$sFilterList = ew_Concat($sFilterList, $this->Date_Updated->AdvancedSearch->ToJSON(), ","); // Field Date_Updated
		$sFilterList = ew_Concat($sFilterList, $this->Updated_By->AdvancedSearch->ToJSON(), ","); // Field Updated_By
		if ($this->BasicSearch->Keyword <> "") {
			$sWrk = "\"" . EW_TABLE_BASIC_SEARCH . "\":\"" . ew_JsEncode2($this->BasicSearch->Keyword) . "\",\"" . EW_TABLE_BASIC_SEARCH_TYPE . "\":\"" . ew_JsEncode2($this->BasicSearch->Type) . "\"";
			$sFilterList = ew_Concat($sFilterList, $sWrk, ",");
		}

		// Return filter list in json
		return ($sFilterList <> "") ? "{" . $sFilterList . "}" : "null";
	}

	// Restore list of filters
	function RestoreFilterList() {
        $restore = new restoreFilterList();
        $restore->RestoreFilterList1($this);
	}

	// Advanced search WHERE clause based on QueryString
	function AdvancedSearchWhere($Default = FALSE) {
		$sWhere = "";
		if (!$_SESSION['Security']->CanSearch()) return "";
		$this->BuildSearchSql($sWhere, $this->Customer_ID, $Default, FALSE); // Customer_ID
		$this->BuildSearchSql($sWhere, $this->Customer_Number, $Default, FALSE); // Customer_Number
		$this->BuildSearchSql($sWhere, $this->Customer_Name, $Default, FALSE); // Customer_Name
		$this->BuildSearchSql($sWhere, $this->Address, $Default, FALSE); // Address
		$this->BuildSearchSql($sWhere, $this->City, $Default, FALSE); // City
		$this->BuildSearchSql($sWhere, $this->Country, $Default, FALSE); // Country
		$this->BuildSearchSql($sWhere, $this->Contact_Person, $Default, FALSE); // Contact_Person
		$this->BuildSearchSql($sWhere, $this->Phone_Number, $Default, FALSE); // Phone_Number
		$this->BuildSearchSql($sWhere, $this->_Email, $Default, FALSE); // Email
		$this->BuildSearchSql($sWhere, $this->Mobile_Number, $Default, FALSE); // Mobile_Number
		$this->BuildSearchSql($sWhere, $this->Notes, $Default, FALSE); // Notes
		$this->BuildSearchSql($sWhere, $this->Balance, $Default, FALSE); // Balance
		$this->BuildSearchSql($sWhere, $this->Date_Added, $Default, FALSE); // Date_Added
		$this->BuildSearchSql($sWhere, $this->Added_By, $Default, FALSE); // Added_By
		$this->BuildSearchSql($sWhere, $this->Date_Updated, $Default, FALSE); // Date_Updated
		$this->BuildSearchSql($sWhere, $this->Updated_By, $Default, FALSE); // Updated_By

		// Set up search parm
		if (!$Default && $sWhere <> "") {
			$this->Command = "search";
		}
		if (!$Default && $this->Command == "search") {
			$this->Customer_ID->AdvancedSearch->Save(); // Customer_ID
			$this->Customer_Number->AdvancedSearch->Save(); // Customer_Number
			$this->Customer_Name->AdvancedSearch->Save(); // Customer_Name
			$this->Address->AdvancedSearch->Save(); // Address
			$this->City->AdvancedSearch->Save(); // City
			$this->Country->AdvancedSearch->Save(); // Country
			$this->Contact_Person->AdvancedSearch->Save(); // Contact_Person
			$this->Phone_Number->AdvancedSearch->Save(); // Phone_Number
			$this->_Email->AdvancedSearch->Save(); // Email
			$this->Mobile_Number->AdvancedSearch->Save(); // Mobile_Number
			$this->Notes->AdvancedSearch->Save(); // Notes
			$this->Balance->AdvancedSearch->Save(); // Balance
			$this->Date_Added->AdvancedSearch->Save(); // Date_Added
			$this->Added_By->AdvancedSearch->Save(); // Added_By
			$this->Date_Updated->AdvancedSearch->Save(); // Date_Updated
			$this->Updated_By->AdvancedSearch->Save(); // Updated_By
		}
		return $sWhere;
	}

	// Build search SQL
	function BuildSearchSql(&$Where, &$Fld, $Default, $MultiValue) {

		$FldVal = ($Default) ? $Fld->AdvancedSearch->SearchValueDefault : $Fld->AdvancedSearch->SearchValue; // @$_GET["x_$FldParm"]
		$FldOpr = ($Default) ? $Fld->AdvancedSearch->SearchOperatorDefault : $Fld->AdvancedSearch->SearchOperator; // @$_GET["z_$FldParm"]
		$FldCond = ($Default) ? $Fld->AdvancedSearch->SearchConditionDefault : $Fld->AdvancedSearch->SearchCondition; // @$_GET["v_$FldParm"]
		$FldVal2 = ($Default) ? $Fld->AdvancedSearch->SearchValue2Default : $Fld->AdvancedSearch->SearchValue2; // @$_GET["y_$FldParm"]
		$FldOpr2 = ($Default) ? $Fld->AdvancedSearch->SearchOperator2Default : $Fld->AdvancedSearch->SearchOperator2; // @$_GET["w_$FldParm"]
		$sWrk = "";

		//$FldVal = ew_StripSlashes($FldVal);
		if (is_array($FldVal)) $FldVal = implode(",", $FldVal);

		//$FldVal2 = ew_StripSlashes($FldVal2);
		if (is_array($FldVal2)) $FldVal2 = implode(",", $FldVal2);
		$FldOpr = strtoupper(trim($FldOpr));
		if ($FldOpr == "") $FldOpr = "=";
		$FldOpr2 = strtoupper(trim($FldOpr2));
		if ($FldOpr2 == "") $FldOpr2 = "=";
		if (EW_SEARCH_MULTI_VALUE_OPTION == 1 || $FldOpr <> "LIKE" ||
			($FldOpr2 <> "LIKE" && $FldVal2 <> ""))
			$MultiValue = FALSE;
		if ($MultiValue) {
			$sWrk1 = ($FldVal <> "") ? ew_GetMultiSearchSql($Fld, $FldOpr, $FldVal, $this->DBID) : ""; // Field value 1
			$sWrk2 = ($FldVal2 <> "") ? ew_GetMultiSearchSql($Fld, $FldOpr2, $FldVal2, $this->DBID) : ""; // Field value 2
			$sWrk = $sWrk1; // Build final SQL
			if ($sWrk2 <> "")
				$sWrk = ($sWrk <> "") ? "($sWrk) $FldCond ($sWrk2)" : $sWrk2;
		} else {
			$FldVal = $this->ConvertSearchValue($Fld, $FldVal);
			$FldVal2 = $this->ConvertSearchValue($Fld, $FldVal2);
			$sWrk = ew_GetSearchSql($Fld, $FldVal, $FldOpr, $FldCond, $FldVal2, $FldOpr2, $this->DBID);
		}
		ew_AddFilter($Where, $sWrk);
	}

	// Convert search value
	function ConvertSearchValue(&$Fld, $FldVal) {
		if ($FldVal == EW_NULL_VALUE || $FldVal == EW_NOT_NULL_VALUE)
			return $FldVal;
		$Value = $FldVal;
		if ($Fld->FldDataType == EW_DATATYPE_BOOLEAN) {
			if ($FldVal <> "") $Value = ($FldVal == "1" || strtolower(strval($FldVal)) == "y" || strtolower(strval($FldVal)) == "t") ? $Fld->TrueValue : $Fld->FalseValue;
		} elseif ($Fld->FldDataType == EW_DATATYPE_DATE) {
			if ($FldVal <> "") $Value = ew_UnFormatDateTime($FldVal, $Fld->FldDateTimeFormat);
		}
		return $Value;
	}

	// Return basic search SQL
	function BasicSearchSQL($arKeywords, $type) {
		$sWhere = "";
		$this->BuildBasicSearchSQL($sWhere, $this->Customer_Number, $arKeywords, $type);
		$this->BuildBasicSearchSQL($sWhere, $this->Customer_Name, $arKeywords, $type);
		$this->BuildBasicSearchSQL($sWhere, $this->Address, $arKeywords, $type);
		$this->BuildBasicSearchSQL($sWhere, $this->City, $arKeywords, $type);
		$this->BuildBasicSearchSQL($sWhere, $this->Country, $arKeywords, $type);
		$this->BuildBasicSearchSQL($sWhere, $this->Contact_Person, $arKeywords, $type);
		$this->BuildBasicSearchSQL($sWhere, $this->Phone_Number, $arKeywords, $type);
		$this->BuildBasicSearchSQL($sWhere, $this->_Email, $arKeywords, $type);
		$this->BuildBasicSearchSQL($sWhere, $this->Mobile_Number, $arKeywords, $type);
		$this->BuildBasicSearchSQL($sWhere, $this->Balance, $arKeywords, $type);
		$this->BuildBasicSearchSQL($sWhere, $this->Added_By, $arKeywords, $type);
		$this->BuildBasicSearchSQL($sWhere, $this->Updated_By, $arKeywords, $type);
		return $sWhere;
	}

	// Build basic search SQL
	function BuildBasicSearchSql(&$Where, &$Fld, $arKeywords, $type) {
		$build= new build_basic_search();
		$build->BuildBasicSearchSql1($Where, $Fld, $arKeywords, $type, $this);
	}

	// Return basic search WHERE clause based on search keyword and type
	function BasicSearchWhere($Default = FALSE) {

		$sSearchStr = "";
		if (!$_SESSION['Security']->CanSearch()) return "";
		$sSearchKeyword = ($Default) ? $this->BasicSearch->KeywordDefault : $this->BasicSearch->Keyword;
		$sSearchType = ($Default) ? $this->BasicSearch->TypeDefault : $this->BasicSearch->Type;
		if ($sSearchKeyword <> "") {
			$sSearch = trim($sSearchKeyword);
			if ($sSearchType <> "=") {
				$ar = array();

				// Match quoted keywords (i.e.: "...")
				if (preg_match_all('/"([^"]*)"/i', $sSearch, $matches, PREG_SET_ORDER)) {
					foreach ($matches as $match) {
						$p = strpos($sSearch, $match[0]);
						$str = substr($sSearch, 0, $p);
						$sSearch = substr($sSearch, $p + strlen($match[0]));
						if (strlen(trim($str)) > 0)
							$ar = array_merge($ar, explode(" ", trim($str)));
						$ar[] = $match[1]; // Save quoted keyword
					}
				}

				// Match individual keywords
				if (strlen(trim($sSearch)) > 0)
					$ar = array_merge($ar, explode(" ", trim($sSearch)));

				// Search keyword in any fields
				if (($sSearchType == "OR" || $sSearchType == "AND") && $this->BasicSearch->BasicSearchAnyFields) {
					foreach ($ar as $sKeyword) {
						if ($sKeyword <> "") {
							if ($sSearchStr <> "") $sSearchStr .= " " . $sSearchType . " ";
							$sSearchStr .= "(" . $this->BasicSearchSQL(array($sKeyword), $sSearchType) . ")";
						}
					}
				} else {
					$sSearchStr = $this->BasicSearchSQL($ar, $sSearchType);
				}
			} else {
				$sSearchStr = $this->BasicSearchSQL(array($sSearch), $sSearchType);
			}
			if (!$Default) $this->Command = "search";
		}
		if (!$Default && $this->Command == "search") {
			$this->BasicSearch->setKeyword($sSearchKeyword);
			$this->BasicSearch->setType($sSearchType);
		}
		return $sSearchStr;
	}

	// Check if search parm exists
	function CheckSearchParms() {

		// Check basic search
		if ($this->BasicSearch->IssetSession())
			return TRUE;
		if ($this->Customer_ID->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->Customer_Number->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->Customer_Name->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->Address->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->City->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->Country->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->Contact_Person->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->Phone_Number->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->_Email->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->Mobile_Number->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->Notes->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->Balance->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->Date_Added->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->Added_By->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->Date_Updated->AdvancedSearch->IssetSession())
			return TRUE;
		if ($this->Updated_By->AdvancedSearch->IssetSession())
			return TRUE;
		return FALSE;
	}

	// Clear all search parameters
	function ResetSearchParms() {

		// Clear search WHERE clause
		$this->SearchWhere = "";
		$this->setSearchWhere($this->SearchWhere);

		// Clear basic search parameters
		$this->ResetBasicSearchParms();

		// Clear advanced search parameters
		$this->ResetAdvancedSearchParms();
	}

	// Load advanced search default values
	function LoadAdvancedSearchDefault() {
		return FALSE;
	}

	// Clear all basic search parameters
	function ResetBasicSearchParms() {
		$this->BasicSearch->UnsetSession();
	}

	// Clear all advanced search parameters
	function ResetAdvancedSearchParms() {
		$this->Customer_ID->AdvancedSearch->UnsetSession();
		$this->Customer_Number->AdvancedSearch->UnsetSession();
		$this->Customer_Name->AdvancedSearch->UnsetSession();
		$this->Address->AdvancedSearch->UnsetSession();
		$this->City->AdvancedSearch->UnsetSession();
		$this->Country->AdvancedSearch->UnsetSession();
		$this->Contact_Person->AdvancedSearch->UnsetSession();
		$this->Phone_Number->AdvancedSearch->UnsetSession();
		$this->_Email->AdvancedSearch->UnsetSession();
		$this->Mobile_Number->AdvancedSearch->UnsetSession();
		$this->Notes->AdvancedSearch->UnsetSession();
		$this->Balance->AdvancedSearch->UnsetSession();
		$this->Date_Added->AdvancedSearch->UnsetSession();
		$this->Added_By->AdvancedSearch->UnsetSession();
		$this->Date_Updated->AdvancedSearch->UnsetSession();
		$this->Updated_By->AdvancedSearch->UnsetSession();
	}

	// Restore all search parameters
	function RestoreSearchParms() {
		$this->RestoreSearch = TRUE;

		// Restore basic search values
		$this->BasicSearch->Load();

		// Restore advanced search values
		$this->Customer_ID->AdvancedSearch->Load();
		$this->Customer_Number->AdvancedSearch->Load();
		$this->Customer_Name->AdvancedSearch->Load();
		$this->Address->AdvancedSearch->Load();
		$this->City->AdvancedSearch->Load();
		$this->Country->AdvancedSearch->Load();
		$this->Contact_Person->AdvancedSearch->Load();
		$this->Phone_Number->AdvancedSearch->Load();
		$this->_Email->AdvancedSearch->Load();
		$this->Mobile_Number->AdvancedSearch->Load();
		$this->Notes->AdvancedSearch->Load();
		$this->Balance->AdvancedSearch->Load();
		$this->Date_Added->AdvancedSearch->Load();
		$this->Added_By->AdvancedSearch->Load();
		$this->Date_Updated->AdvancedSearch->Load();
		$this->Updated_By->AdvancedSearch->Load();
	}

	// Set up sort parameters
	function SetUpSortOrder() {

		// Check for "order" parameter
		if ($_GET["order"] <> "") {
			$this->CurrentOrder = ew_StripSlashes($_GET["order"]);
			$this->CurrentOrderType = $_GET["ordertype"];
			$this->UpdateSort($this->Customer_Number); // Customer_Number
			$this->UpdateSort($this->Customer_Name); // Customer_Name
			$this->UpdateSort($this->Contact_Person); // Contact_Person
			$this->UpdateSort($this->Phone_Number); // Phone_Number
			$this->UpdateSort($this->Mobile_Number); // Mobile_Number
			$this->UpdateSort($this->Balance); // Balance
			$this->setStartRecordNumber(1); // Reset start position
		}
	}

	// Load sort order parameters
	function LoadSortOrder() {
		$sOrderBy = $this->getSessionOrderBy(); // Get ORDER BY from Session
		if ($sOrderBy == "") {
			if ($this->getSqlOrderBy() <> "") {
				$sOrderBy = $this->getSqlOrderBy();
				$this->setSessionOrderBy($sOrderBy);
			}
		}
	}

	// Reset command
	// - cmd=reset (Reset search parameters)
	// - cmd=resetall (Reset search and master/detail parameters)
	// - cmd=resetsort (Reset sort parameters)
	function ResetCmd() {

		// Check if reset command
		if (substr($this->Command,0,5) == "reset") {

			// Reset search criteria
			if ($this->Command == "reset" || $this->Command == "resetall")
				$this->ResetSearchParms();

			// Reset sorting order
			if ($this->Command == "resetsort") {
				$sOrderBy = "";
				$this->setSessionOrderBy($sOrderBy);
				$this->Customer_Number->setSort("");
				$this->Customer_Name->setSort("");
				$this->Contact_Person->setSort("");
				$this->Phone_Number->setSort("");
				$this->Mobile_Number->setSort("");
				$this->Balance->setSort("");
			}

			// Reset start position
			$this->StartRec = 1;
			$this->setStartRecordNumber($this->StartRec);
		}
	}

	// Set up list options
	function SetupListOptions() {


		// Add group option item
		$item = &$this->ListOptions->Add($this->ListOptions->GroupOptionName);
		$item->Body = "";
		$item->OnLeft = TRUE;
		$item->Visible = FALSE;

		// "view"
		$item = &$this->ListOptions->Add("view");
		$item->CssStyle = "white-space: nowrap;";
		$item->Visible = $_SESSION['Security']->CanView();
		$item->OnLeft = TRUE;

		// "edit"
		$item = &$this->ListOptions->Add("edit");
		$item->CssStyle = "white-space: nowrap;";
		$item->Visible = $_SESSION['Security']->CanEdit();
		$item->OnLeft = TRUE;

		// "copy"
		$item = &$this->ListOptions->Add("copy");
		$item->CssStyle = "white-space: nowrap;";
		$item->Visible = $_SESSION['Security']->CanAdd();
		$item->OnLeft = TRUE;

		// "detail_a_sales"
		$item = &$this->ListOptions->Add("detail_a_sales");
		$item->CssStyle = "white-space: nowrap;";
		$item->Visible = $_SESSION['Security']->AllowList(CurrentProjectID() . 'a_sales') && !$this->ShowMultipleDetails;
		$item->OnLeft = TRUE;
		$item->ShowInButtonGroup = FALSE;
		if (!isset($GLOBALS["a_sales_grid"])) $GLOBALS["a_sales_grid"] = new ca_sales_grid;

		// Multiple details
		if ($this->ShowMultipleDetails) {
			$item = &$this->ListOptions->Add("details");
			$item->CssStyle = "white-space: nowrap;";
			$item->Visible = $this->ShowMultipleDetails;
			$item->OnLeft = TRUE;
			$item->ShowInButtonGroup = FALSE;
		}

		// Set up detail pages
		$pages = new cSubPages();
		$pages->Add("a_sales");
		$this->DetailPages = $pages;

		// List actions
		$item = &$this->ListOptions->Add("listactions");
		$item->CssStyle = "white-space: nowrap;";
		$item->OnLeft = TRUE;
		$item->Visible = FALSE;
		$item->ShowInButtonGroup = FALSE;
		$item->ShowInDropDown = FALSE;

		// "checkbox"
		$item = &$this->ListOptions->Add("checkbox");
		$item->Visible = ($_SESSION['Security']->CanDelete() || $_SESSION['Security']->CanEdit());
		$item->OnLeft = TRUE;
		$item->Header = "<input type=\"checkbox\" name=\"key\" id=\"key\" onclick=\"ew_SelectAllKey(this);\">";
		$item->MoveTo(0);
		$item->ShowInDropDown = FALSE;
		$item->ShowInButtonGroup = FALSE;

		// Drop down button for ListOptions
		$this->ListOptions->UseImageAndText = TRUE;
		$this->ListOptions->UseDropDownButton = TRUE;
		$this->ListOptions->DropDownButtonPhrase = $_SESSION['Language']->Phrase("ButtonListOptions");
		$this->ListOptions->UseButtonGroup = FALSE;
		if ($this->ListOptions->UseButtonGroup && ew_IsMobile())
			$this->ListOptions->UseDropDownButton = TRUE;
		$this->ListOptions->ButtonClass = "btn-sm"; // Class for button group

		// Call ListOptions_Load event
		$this->ListOptions_Load();
		$this->SetupListOptionsExt();
		$item = &$this->ListOptions->GetItem($this->ListOptions->GroupOptionName);
		$item->Visible = $this->ListOptions->GroupOptionVisible();
	}

	// Render list options
	function RenderListOptions() {
		$listOption= new render_list();
		$listOption->RenderListOptions1($this);
	}

	// Set up other options
	function SetupOtherOptions() {

		$setup= new setup_option();
		$setup->SetupOtherOption1($this);
	}

	// Render other options
	function RenderOtherOptions() {

		$options = &$this->OtherOptions;
			$option = &$options["action"];

			// Set up list action buttons
			foreach ($this->ListActions->Items as $listaction) {
				if ($listaction->Select == EW_ACTION_MULTIPLE) {
					$item = &$option->Add("custom_" . $listaction->Action);
					$caption = $listaction->Caption;
					$icon = ($listaction->Icon <> "") ? "<span class=\"" . ew_HtmlEncode($listaction->Icon) . "\" data-caption=\"" . ew_HtmlEncode($caption) . "\"></span> " : $caption;
					$item->Body = "<a class=\"ewAction ewListAction\" title=\"" . ew_HtmlEncode($caption) . "\" data-caption=\"" . ew_HtmlEncode($caption) . "\" href=\"\" onclick=\"ew_SubmitAction(event,jQuery.extend({f:document.fa_customerslist}," . $listaction->ToJson(TRUE) . "));return false;\">" . $icon . "</a>";
					$item->Visible = $listaction->Allow;
				}
			}

			// Hide grid edit and other options
			if ($this->TotalRecs <= 0) {
				$option = &$options["addedit"];
				$item = &$option->GetItem("gridedit");
				if ($item) $item->Visible = FALSE;
				$option = &$options["action"];
				$option->HideAllOptions();
			}
	}

	// Process list action
	function ProcessListAction() {
	    $process= new list_action();
	    return $process->ProcessListAction1($this);
	}

	// Set up search options
	function SetupSearchOptions() {

		$this->SearchOptions = new cListOptions();
		$this->SearchOptions->Tag = "div";
		$this->SearchOptions->TagClassName = "ewSearchOption";

		// Search button
		$item = &$this->SearchOptions->Add("searchtoggle");

		// Begin of modification Customizing Search Panel, by Masino Sinaga, for customize search panel, July 22, 2014
		if (MS_USE_TABLE_SETTING_FOR_SEARCH_PANEL_COLLAPSED) {			

			// The code in this first block will be generated if "UseTableSettingForSearchPanelCollapsed" is enabled from "MasinoFixedWidthSite12" extension, also with "InitSearchPanelAsCollapsed" is enabled from -> "Advanced" -> "Tables" setting.
			if ($this->SearchPanelCollapsed==TRUE) {
				$SearchToggleClass = " ";
			} else {
				$SearchToggleClass = " active";
			}
		} else {

			// Nothing to do, because we've been using MS_SEARCH_PANEL_COLLAPSED value from the generated "ewcfg11.php" file
			// $SearchToggleClass = ($this->SearchWhere <> "") ? " active" : " active"; // <-- no need to use this anymore!

			if (MS_SEARCH_PANEL_COLLAPSED == TRUE && $this->SearchWhere <> "") {
				$SearchToggleClass = " active";
			} elseif (MS_SEARCH_PANEL_COLLAPSED == TRUE && $this->SearchWhere == "") {
				$SearchToggleClass = " ";
			} elseif (MS_SEARCH_PANEL_COLLAPSED == FALSE && $this->SearchWhere <> "") {
				$SearchToggleClass = " active";			
			} elseif (MS_SEARCH_PANEL_COLLAPSED == FALSE && $this->SearchWhere == "") {
				$SearchToggleClass = " active";
			}
		}

		// End of modification Customizing Search Panel, by Masino Sinaga, for customize search panel, July 22, 2014
		// Begin of modification Hide Search Button for Inline Edit and Inline Copy mode in List Page, by Masino Sinaga, August 4, 2014

		if ($this->CurrentAction == "edit" || $this->CurrentAction == "copy") {
		} else {
			$item->Body = "<button type=\"button\" class=\"btn btn-default ewSearchToggle" . $SearchToggleClass . "\" title=\"" . $_SESSION['Language']->Phrase("SearchPanel") . "\" data-caption=\"" . $_SESSION['Language']->Phrase("SearchPanel") . "\" data-toggle=\"button\" data-form=\"fa_customerslistsrch\">" . $_SESSION['Language']->Phrase("SearchBtn") . "</button>";
			$item->Visible = TRUE;
		}

		// End of modification Hide Search Button for Inline Edit and Inline Copy mode in List Page, by Masino Sinaga, August 4, 2014			
		// Begin of modification Hide Search Button for Inline Edit and Inline Copy mode in List Page, by Masino Sinaga, August 4, 2014

		if ($this->CurrentAction == "edit" || $this->CurrentAction == "copy") {
		} else {

			// Show all button
			$item = &$this->SearchOptions->Add("showall");
			$item->Body = "<a class=\"btn btn-default ewShowAll\" title=\"" . $_SESSION['Language']->Phrase("ShowAll") . "\" data-caption=\"" . $_SESSION['Language']->Phrase("ShowAll") . "\" href=\"" . $this->PageUrl() . "cmd=reset\">" . $_SESSION['Language']->Phrase("ShowAllBtn") . "</a>";
			$item->Visible = ($this->SearchWhere <> $this->DefaultSearchWhere && $this->SearchWhere <> "0=101"); // v11.0.4
		}

		// End of modification Hide Search Button for Inline Edit and Inline Copy mode in List Page, by Masino Sinaga, August 4, 2014
		// Advanced search button

		$item = &$this->SearchOptions->Add("advancedsearch");
		$item->Body = "<a class=\"btn btn-default ewAdvancedSearch\" title=\"" . $_SESSION['Language']->Phrase("AdvancedSearch") . "\" data-caption=\"" . $_SESSION['Language']->Phrase("AdvancedSearch") . "\" href=\"a_customerssrch.php\">" . $_SESSION['Language']->Phrase("AdvancedSearchBtn") . "</a>";
		$item->Visible = TRUE;

		// Button group for search
		$this->SearchOptions->UseDropDownButton = FALSE;
		$this->SearchOptions->UseImageAndText = TRUE;
		$this->SearchOptions->UseButtonGroup = TRUE;
		$this->SearchOptions->DropDownButtonPhrase = $_SESSION['Language']->Phrase("ButtonSearch");

		// Add group option item
		$item = &$this->SearchOptions->Add($this->SearchOptions->GroupOptionName);
		$item->Body = "";
		$item->Visible = FALSE;

		// Hide search options
		if ($this->Export <> "" || $this->CurrentAction <> "")
			$this->SearchOptions->HideAllOptions();

		if (!$_SESSION['Security']->CanSearch()) {
			$this->SearchOptions->HideAllOptions();
			$this->FilterOptions->HideAllOptions();
		}
	}

	function SetupListOptionsExt() {

	}

	function RenderListOptionsExt() {
		$ext= new render_option_ext();
		$ext->RenderListOptionsExt1($this);
	}

	// Set up starting record parameters
	function SetUpStartRec() {
        $start= new funct_2();
        $start->SetUpStartRec1($this);
	}

	// Load basic search values
	function LoadBasicSearchValues() {
		$this->BasicSearch->Keyword = $_GET[EW_TABLE_BASIC_SEARCH];
		if ($this->BasicSearch->Keyword <> "") $this->Command = "search";
		$this->BasicSearch->Type = $_GET[EW_TABLE_BASIC_SEARCH_TYPE];
	}

	// Load search values for validation
	function LoadSearchValues() {
		$load=new search2();
		$load->LoadSearchValues2($this);
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
		$this->ViewUrl = $this->GetViewUrl();
		$this->EditUrl = $this->GetEditUrl();
		$this->InlineEditUrl = $this->GetInlineEditUrl();
		$this->CopyUrl = $this->GetCopyUrl();
		$this->InlineCopyUrl = $this->GetInlineCopyUrl();
		$this->DeleteUrl = $this->GetDeleteUrl();

		// Convert decimal values if posted back
		if ($this->Balance->FormValue == $this->Balance->CurrentValue && is_numeric(ew_StrToFloat($this->Balance->CurrentValue)))
			$this->Balance->CurrentValue = ew_StrToFloat($this->Balance->CurrentValue);

		// Call Row_Rendering event
		$this->Row_Rendering();

		// Common render codes for all row types
		// Customer_ID

		$this->Customer_ID->CellCssStyle = "white-space: nowrap;";

		// Customer_Number
		// Customer_Name
		// Address

		$this->Address->CellCssStyle = "white-space: nowrap;";

		// City
		$this->City->CellCssStyle = "white-space: nowrap;";

		// Country
		$this->Country->CellCssStyle = "white-space: nowrap;";

		// Contact_Person
		// Phone_Number
		// Email
		// Mobile_Number
		// Notes

		$this->Notes->CellCssStyle = "white-space: nowrap;";

		// Balance
		// Date_Added
		// Added_By
		// Date_Updated
		// Updated_By
		// Accumulate aggregate value

		if ($this->RowType <> EW_ROWTYPE_AGGREGATEINIT && $this->RowType <> EW_ROWTYPE_AGGREGATE) {
			if (is_numeric($this->Balance->CurrentValue))
				$this->Balance->Total += $this->Balance->CurrentValue; // Accumulate total
		}
		if ($this->RowType == EW_ROWTYPE_VIEW) { // View row

		// Customer_Number
		$this->Customer_Number->ViewValue = $this->Customer_Number->CurrentValue;
		$this->Customer_Number->ViewCustomAttributes = "";

		// Customer_Name
		$this->Customer_Name->ViewValue = $this->Customer_Name->CurrentValue;
		$this->Customer_Name->ViewCustomAttributes = "";

		// Contact_Person
		$this->Contact_Person->ViewValue = $this->Contact_Person->CurrentValue;
		$this->Contact_Person->ViewCustomAttributes = "";

		// Phone_Number
		$this->Phone_Number->ViewValue = $this->Phone_Number->CurrentValue;
		$this->Phone_Number->ViewCustomAttributes = "";

		// Mobile_Number
		$this->Mobile_Number->ViewValue = $this->Mobile_Number->CurrentValue;
		$this->Mobile_Number->ViewCustomAttributes = "";

		// Balance
		$this->Balance->ViewValue = $this->Balance->CurrentValue;
		$this->Balance->ViewValue = ew_FormatCurrency($this->Balance->ViewValue, 2, -2, -2, -2);
		$this->Balance->CellCssStyle .= "text-align: right;";
		$this->Balance->ViewCustomAttributes = "";

			// Customer_Number
			$this->Customer_Number->LinkCustomAttributes = "";
			$this->Customer_Number->HrefValue = "";
			$this->Customer_Number->TooltipValue = "";

			// Customer_Name
			$this->Customer_Name->LinkCustomAttributes = "";
			$this->Customer_Name->HrefValue = "";
			$this->Customer_Name->TooltipValue = "";

			// Contact_Person
			$this->Contact_Person->LinkCustomAttributes = "";
			$this->Contact_Person->HrefValue = "";
			$this->Contact_Person->TooltipValue = "";

			// Phone_Number
			$this->Phone_Number->LinkCustomAttributes = "";
			$this->Phone_Number->HrefValue = "";
			$this->Phone_Number->TooltipValue = "";

			// Mobile_Number
			$this->Mobile_Number->LinkCustomAttributes = "";
			$this->Mobile_Number->HrefValue = "";
			$this->Mobile_Number->TooltipValue = "";

			// Balance
			$this->Balance->LinkCustomAttributes = "";
			$this->Balance->HrefValue = "";
			$this->Balance->TooltipValue = "";
		} elseif ($this->RowType == EW_ROWTYPE_AGGREGATEINIT) { // Initialize aggregate row
			$this->Balance->Total = 0; // Initialize total
		} elseif ($this->RowType == EW_ROWTYPE_AGGREGATE) { // Aggregate row
			$this->Balance->CurrentValue = $this->Balance->Total;
			$this->Balance->ViewValue = $this->Balance->CurrentValue;
			$this->Balance->ViewValue = ew_FormatCurrency($this->Balance->ViewValue, 2, -2, -2, -2);
			$this->Balance->CellCssStyle .= "text-align: right;";
			$this->Balance->ViewCustomAttributes = "";
			$this->Balance->HrefValue = ""; // Clear href value
		}

		// Call Row Rendered event
		if ($this->RowType <> EW_ROWTYPE_AGGREGATEINIT)
			$this->Row_Rendered();
	}

	// Validate search
	function ValidateSearch() {
		

		// Initialize
		$_SESSION['$gsSearchError'] = "";

		// Check if validation required
		if (!EW_SERVER_VALIDATE)
			return TRUE;

		// Return validate result
		$ValidateSearch = ($_SESSION['$gsSearchError'] == "");

		// Call Form_CustomValidate event
		$sFormCustomError = "";
		$ValidateSearch = $ValidateSearch && $this->Form_CustomValidate();
		if ($sFormCustomError <> "") {
			ew_AddMessage($_SESSION['$gsSearchError'], $sFormCustomError);
		}
		return $ValidateSearch;
	}

	// Load advanced search
	function LoadAdvancedSearch() {
		$this->Customer_ID->AdvancedSearch->Load();
		$this->Customer_Number->AdvancedSearch->Load();
		$this->Customer_Name->AdvancedSearch->Load();
		$this->Address->AdvancedSearch->Load();
		$this->City->AdvancedSearch->Load();
		$this->Country->AdvancedSearch->Load();
		$this->Contact_Person->AdvancedSearch->Load();
		$this->Phone_Number->AdvancedSearch->Load();
		$this->_Email->AdvancedSearch->Load();
		$this->Mobile_Number->AdvancedSearch->Load();
		$this->Notes->AdvancedSearch->Load();
		$this->Balance->AdvancedSearch->Load();
		$this->Date_Added->AdvancedSearch->Load();
		$this->Added_By->AdvancedSearch->Load();
		$this->Date_Updated->AdvancedSearch->Load();
		$this->Updated_By->AdvancedSearch->Load();
	}

	// Build export filter for selected records
	function BuildExportSelectedFilter() {

		$sWrkFilter = "";
		if ($this->Export <> "") {
			$sWrkFilter = $this->GetKeyFilter();
		}
		return $sWrkFilter;
	}

	// Set up export options
	function SetupExportOptions() {

// Begin of modification Permission Access for Export To Feature, by Masino Sinaga, May 5, 2012
        

		// Printer friendly
        if ($_SESSION['Security']->CanExportToPrint() || $_SESSION['Security']->IsAdmin() ) {
			$item = &$this->ExportOptions->Add("print");

			// $item->Body = "<a href=\"" . $this->ExportPrintUrl . "\" class=\"ewExportLink ewPrint\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("PrinterFriendlyText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("PrinterFriendlyText")) . "\">" . $_SESSION['Language']->Phrase("PrinterFriendly") . "</a>";
			// Begin of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012

			if (MS_EXPORT_RECORD_OPTIONS=="selectedrecords") {
				$item->Body = "<a class=\"ewExportLink ewPrint\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("PrinterFriendlyText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("PrinterFriendlyText")) . "\" onclick=\"ew_Export(document.fa_customerslist,'" . ew_CurrentPage() . "','print',false,true);\">" . $_SESSION['Language']->Phrase("PrinterFriendly") . "</a>";
			} else {
				$item->Body = "<a href=\"" . $this->ExportPrintUrl . "\" class=\"ewExportLink ewPrint\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("PrinterFriendlyText")) . "\"  data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("PrinterFriendlyText")) . "\">" . $_SESSION['Language']->Phrase("PrinterFriendly") . "</a>";
			}

			// End of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012
			$item->Visible = TRUE;
        }

		// Export to Excel
        if ($_SESSION['Security']->CanExportToExcel() || $_SESSION['Security']->IsAdmin() ) {
			$item = &$this->ExportOptions->Add("excel");

			// $item->Body = "<a href=\"" . $this->ExportExcelUrl . "\" class=\"ewExportLink ewExcel\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToExcelText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToExcelText")) . "\">" . $_SESSION['Language']->Phrase("ExportToExcel") . "</a>";
			// Begin of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012

			if (MS_EXPORT_RECORD_OPTIONS=="selectedrecords") {
				$item->Body = "<a class=\"ewExportLink ewExcel\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToExcelText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToExcelText")) . "\" onclick=\"ew_Export(document.fa_customerslist,'" . ew_CurrentPage() . "','excel',false,true);\">" . $_SESSION['Language']->Phrase("ExportToExcel") . "</a>";
			} else {
				$item->Body = "<a href=\"" . $this->ExportExcelUrl . "\" class=\"ewExportLink ewExcel\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToExcelText")) . "\"  data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToExcelText")) . "\">" . $_SESSION['Language']->Phrase("ExportToExcel") . "</a>";
			}

			// End of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012
			$item->Visible = TRUE;
        }

		// Export to Word
        if ($_SESSION['Security']->CanExportToWord() || $_SESSION['Security']->IsAdmin() ) {
			$item = &$this->ExportOptions->Add("word");

			// $item->Body = "<a href=\"" . $this->ExportWordUrl . "\" class=\"ewExportLink ewWord\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToWordText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToWordText")) . "\">" . $_SESSION['Language']->Phrase("ExportToWord") . "</a>";
			// Begin of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012

			if (MS_EXPORT_RECORD_OPTIONS=="selectedrecords") {
				$item->Body = "<a class=\"ewExportLink ewWord\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToWordText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToWordText")) . "\" onclick=\"ew_Export(document.fa_customerslist,'" . ew_CurrentPage() . "','word',false,true);\">" . $_SESSION['Language']->Phrase("ExportToWord") . "</a>";
			} else {
				$item->Body = "<a href=\"" . $this->ExportWordUrl . "\" class=\"ewExportLink ewWord\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToWordText")) . "\"  data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToWordText")) . "\">" . $_SESSION['Language']->Phrase("ExportToWord") . "</a>";
			}

			// End of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012
			$item->Visible = TRUE;
        }

		// Export to Html
        if ($_SESSION['Security']->CanExportToHTML() || $_SESSION['Security']->IsAdmin() ) {
			$item = &$this->ExportOptions->Add("html");

			// $item->Body = "<a href=\"" . $this->ExportHtmlUrl . "\" class=\"ewExportLink ewHtml\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToHtmlText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToHtmlText")) . "\">" . $_SESSION['Language']->Phrase("ExportToHtml") . "</a>";
			// Begin of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012

			if (MS_EXPORT_RECORD_OPTIONS=="selectedrecords") {
				$item->Body = "<a class=\"ewExportLink ewHtml\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToHtmlText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToHtmlText")) . "\" onclick=\"ew_Export(document.fa_customerslist,'" . ew_CurrentPage() . "','html',false,true);\">" . $_SESSION['Language']->Phrase("ExportToHtml") . "</a>";
			} else {
				$item->Body = "<a href=\"" . $this->ExportHtmlUrl . "\" class=\"ewExportLink ewHtml\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToHtmlText")) . "\"  data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToHtmlText")) . "\">" . $_SESSION['Language']->Phrase("ExportToHTML") . "</a>";
			}

			// End of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012
			$item->Visible = TRUE;
        }

		// Export to Xml
        if ($_SESSION['Security']->CanExportToXML() || $_SESSION['Security']->IsAdmin() ) {
			$item = &$this->ExportOptions->Add("xml");

			// $item->Body = "<a href=\"" . $this->ExportXmlUrl . "\" class=\"ewExportLink ewXml\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToXmlText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToXmlText")) . "\">" . $_SESSION['Language']->Phrase("ExportToXml") . "</a>";
			// Begin of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012

			if (MS_EXPORT_RECORD_OPTIONS=="selectedrecords") {
				$item->Body = "<a class=\"ewExportLink ewXml\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToXmlText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToXmlText")) . "\" onclick=\"ew_Export(document.fa_customerslist,'" . ew_CurrentPage() . "','xml',false,true);\">" . $_SESSION['Language']->Phrase("ExportToXml") . "</a>";
			} else {
				$item->Body = "<a href=\"" . $this->ExportXmlUrl . "\" class=\"ewExportLink ewXml\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToXmlText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToXmlText")) . "\">" . $_SESSION['Language']->Phrase("ExportToXML") . "</a>";
			}

			// End of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012
			$item->Visible = TRUE;
        }

		// Export to Csv
        if ($_SESSION['Security']->CanExportToCSV() || $_SESSION['Security']->IsAdmin() ) {
			$item = &$this->ExportOptions->Add("csv");

			// $item->Body = "<a href=\"" . $this->ExportCsvUrl . "\" class=\"ewExportLink ewCsv\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToCsvText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToCsvText")) . "\">" . $_SESSION['Language']->Phrase("ExportToCsv") . "</a>";
			// Begin of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012

			if (MS_EXPORT_RECORD_OPTIONS=="selectedrecords") {
				$item->Body = "<a class=\"ewExportLink ewCsv\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToCsvText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToCsvText")) . "\" onclick=\"ew_Export(document.fa_customerslist,'" . ew_CurrentPage() . "','csv',false,true);\">" . $_SESSION['Language']->Phrase("ExportToCsv") . "</a>";
			} else {
				$item->Body = "<a href=\"" . $this->ExportCsvUrl . "\" class=\"ewExportLink ewCsv\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToCsvText")) . "\"  data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToCsvText")) . "\">" . $_SESSION['Language']->Phrase("ExportToCSV") . "</a>";
			}

			// End of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012
			$item->Visible = TRUE;
        }

		// Export to Pdf
        if ($_SESSION['Security']->CanExportToPDF() || $_SESSION['Security']->IsAdmin() ) {
			$item = &$this->ExportOptions->Add("pdf");

			// $item->Body = "<a href=\"" . $this->ExportPdfUrl . "\" class=\"ewExportLink ewPdf\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToPDFText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToPDFText")) . "\">" . $_SESSION['Language']->Phrase("ExportToPDF") . "</a>";
			// Begin of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012

			if (MS_EXPORT_RECORD_OPTIONS=="selectedrecords") {
				$item->Body = "<a class=\"ewExportLink ewPdf\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToPDFText")) . "\" data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToPDFText")) . "\" onclick=\"ew_Export(document.fa_customerslist,'" . ew_CurrentPage() . "','pdf',false,true);\">" . $_SESSION['Language']->Phrase("ExportToPDF") . "</a>";
			} else {
				$item->Body = "<a href=\"" . $this->ExportPdfUrl . "\" class=\"ewExportLink ewPdf\" title=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToPDFText")) . "\"  data-caption=\"" . ew_HtmlEncode($_SESSION['Language']->Phrase("ExportToPDFText")) . "\">" . $_SESSION['Language']->Phrase("ExportToPDF") . "</a>";
			}

			// End of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012
			$item->Visible = FALSE;
        }

		// Export to Email
		if ($_SESSION['Security']->CanExportToEmail() || $_SESSION['Security']->IsAdmin() ) {
			$item = &$this->ExportOptions->Add("email");


			// $item->Body = "<button id=\"emf_a_customers\" class=\"ewExportLink ewEmail\" title=\"" . $_SESSION['Language']->Phrase("ExportToEmailText") . "\" data-caption=\"" . $_SESSION['Language']->Phrase("ExportToEmailText") . "\" onclick=\"ew_EmailDialogShow({lnk:'emf_a_customers',hdr:ewLanguage.Phrase('ExportToEmailText'),f:document.fa_customerslist,sel:false" . $url . "});\">" . $_SESSION['Language']->Phrase("ExportToEmail") . "</button>";
			// Begin of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012

		if (MS_EXPORT_RECORD_OPTIONS=="selectedrecords") {
			$item->Body = "<a id=\"emf_a_customers\" href=\"javascript:void(0);\" class=\"ewExportLink ewEmail\" title=\"" . $_SESSION['Language']->Phrase("ExportToEmailText") . "\"  data-caption=\"" . $_SESSION['Language']->Phrase("ExportToEmailText") . "\" onclick=\"ew_EmailDialogShow({lnk:'emf_a_customers',hdr:ewLanguage.Phrase('ExportToEmailText'),f:document.fa_customerslist,sel:true});\">" . $_SESSION['Language']->Phrase("ExportToEmail") . "</a>";
		} else {
			$item->Body = "<a id=\"emf_a_customers\" href=\"javascript:void(0);\" class=\"ewExportLink ewEmail\" title=\"" . $_SESSION['Language']->Phrase("ExportToEmailText") . "\"  data-caption=\"" . $_SESSION['Language']->Phrase("ExportToEmailText") . "\" onclick=\"ew_EmailDialogShow({lnk:'emf_a_customers',hdr:ewLanguage.Phrase('ExportToEmailText'),f:document.fa_customerslist,sel:false});\">" . $_SESSION['Language']->Phrase("ExportToEmail") . "</a>";
		}

		// End of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012
			$item->Visible = TRUE;
        }

		// Drop down button for export
		$this->ExportOptions->UseButtonGroup = TRUE;
		$this->ExportOptions->UseImageAndText = TRUE;
		$this->ExportOptions->UseDropDownButton = TRUE;
		if ($this->ExportOptions->UseButtonGroup && ew_IsMobile())
			$this->ExportOptions->UseDropDownButton = TRUE;
		$this->ExportOptions->DropDownButtonPhrase = $_SESSION['Language']->Phrase("ButtonExport");

		// Add group option item
		$item = &$this->ExportOptions->Add($this->ExportOptions->GroupOptionName);
		$item->Body = "";
		$item->Visible = FALSE;
	}

	// Export data in HTML/CSV/Word/Excel/XML/Email/PDF format
	function ExportData() {

		$bSelectLimit = $this->UseSelectLimit;

		// Load recordset
		if ($bSelectLimit) {
			$this->TotalRecs = $this->SelectRecordCount();
		} else {

			// changed since v11.0.6
			if (!$this->Recordset)
				$this->Recordset = $this->LoadRecordset();
			$rs = &$this->Recordset;
			if ($rs)
				$this->TotalRecs = $rs->RecordCount();
		}
		$this->StartRec = 1;

		// Export all
		// Begin of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012

		if ($this->ExportAll=="allpages") {

		// End of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012
			set_time_limit(EW_EXPORT_ALL_TIME_LIMIT);
			$this->DisplayRecs = $this->TotalRecs;
			$this->StopRec = $this->TotalRecs;
		} else { // Export one page only
			$this->SetUpStartRec(); // Set up start record position

			// Set the last record to display
			if ($this->DisplayRecs <= 0) {
				$this->StopRec = $this->TotalRecs;
			} else {
				$this->StopRec = $this->StartRec + $this->DisplayRecs - 1;
			}
		}
		if ($bSelectLimit)
			$rs = $this->LoadRecordset();
		if (!$rs) {
			header("Content-Type:"); // Remove header
			header("Content-Disposition:");
			$this->ShowMessage();
			return;
		}
		$this->ExportDoc = ew_ExportDocument($this, "h");
		$Doc = &$this->ExportDoc;
		if ($bSelectLimit) {
			$this->StartRec = 1;
			$this->StopRec = $this->DisplayRecs <= 0 ? $this->TotalRecs : $this->DisplayRecs;
		} else {

			//$this->StartRec = $this->StartRec;
			//$this->StopRec = $this->StopRec;

		}

		// Call Page Exporting server event
		$this->ExportDoc->ExportCustom = !$this->Page_Exporting();

		$sHeader = $this->PageHeader;
		$this->Page_DataRendering();
		$Doc->Text .= $sHeader;
		$this->ExportDocument($Doc, $rs, $this->StartRec, $this->StopRec, "");
		$sFooter = $this->PageFooter;
		$this->Page_DataRendered();
		$Doc->Text .= $sFooter;

		// Close recordset
		$rs->Close();

		// Call Page Exported server event
		$this->Page_Exported();

		// Export header and footer
		$Doc->ExportHeaderAndFooter();

		// Clean output buffer
		if (!EW_DEBUG_ENABLED && ob_get_length())
			ob_end_clean();

		// Write debug message if enabled
		if (EW_DEBUG_ENABLED && $this->Export <> "pdf")
			echo ew_DebugMsg();

		// Output data
		if ($this->Export == "email") {
			echo $this->ExportEmail($Doc->Text);
		} else {
			$Doc->Export();
		}
	}

	// Export email
	function ExportEmail($EmailContent){
	    $expem= new export_email();
	    $expem->ExportEmail1($EmailContent, $this);
    }

	// Export QueryString
	function ExportQueryString() {

		// Initialize
		$sQry = "export=html";

		// Build QueryString for search
		if ($this->BasicSearch->getKeyword() <> "") {
			$sQry .= "&" . EW_TABLE_BASIC_SEARCH . "=" . urlencode($this->BasicSearch->getKeyword()) . "&" . EW_TABLE_BASIC_SEARCH_TYPE . "=" . urlencode($this->BasicSearch->getType());
		}
		$this->AddSearchQueryString($sQry, $this->Customer_ID); // Customer_ID
		$this->AddSearchQueryString($sQry, $this->Customer_Number); // Customer_Number
		$this->AddSearchQueryString($sQry, $this->Customer_Name); // Customer_Name
		$this->AddSearchQueryString($sQry, $this->Address); // Address
		$this->AddSearchQueryString($sQry, $this->City); // City
		$this->AddSearchQueryString($sQry, $this->Country); // Country
		$this->AddSearchQueryString($sQry, $this->Contact_Person); // Contact_Person
		$this->AddSearchQueryString($sQry, $this->Phone_Number); // Phone_Number
		$this->AddSearchQueryString($sQry, $this->_Email); // Email
		$this->AddSearchQueryString($sQry, $this->Mobile_Number); // Mobile_Number
		$this->AddSearchQueryString($sQry, $this->Notes); // Notes
		$this->AddSearchQueryString($sQry, $this->Balance); // Balance
		$this->AddSearchQueryString($sQry, $this->Date_Added); // Date_Added
		$this->AddSearchQueryString($sQry, $this->Added_By); // Added_By
		$this->AddSearchQueryString($sQry, $this->Date_Updated); // Date_Updated
		$this->AddSearchQueryString($sQry, $this->Updated_By); // Updated_By

		// Build QueryString for pager
		$sQry .= "&" . EW_TABLE_REC_PER_PAGE . "=" . urlencode($this->getRecordsPerPage()) . "&" . EW_TABLE_START_REC . "=" . urlencode($this->getStartRecordNumber());
		return $sQry;
	}

	// Add search QueryString
	function AddSearchQueryString(&$Qry, &$Fld) {
		$FldSearchValue = $Fld->AdvancedSearch->getValue("x");
		$FldParm = substr($Fld->FldVar,2);
		if (strval($FldSearchValue) <> "") {
			$Qry .= "&x_" . $FldParm . "=" . urlencode($FldSearchValue) .
				"&z_" . $FldParm . "=" . urlencode($Fld->AdvancedSearch->getValue("z"));
		}
		$FldSearchValue2 = $Fld->AdvancedSearch->getValue("y");
		if (strval($FldSearchValue2) <> "") {
			$Qry .= "&v_" . $FldParm . "=" . urlencode($Fld->AdvancedSearch->getValue("v")) .
				"&y_" . $FldParm . "=" . urlencode($FldSearchValue2) .
				"&w_" . $FldParm . "=" . urlencode($Fld->AdvancedSearch->getValue("w"));
		}
	}

	// Set up Breadcrumb
	function SetupBreadcrumb() {

		$_SESSION['Breadcrumb'] = new cBreadcrumb();
		$url = substr(ew_CurrentUrl(), strrpos(ew_CurrentUrl(), "/")+1); // v11.0.4

		// $url = ew_CurrentUrl(); // <-- removed since v11.0.4
		$url = preg_replace('/\?cmd=reset(all){0,1}$/i', '', $url); // Remove cmd=reset / cmd=resetall
        $_SESSION['Breadcrumb']->Add("list", $this->TableVar, $url, "", $this->TableVar, TRUE);
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
		$this->ListOptions->UseDropDownButton = FALSE;
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

	// ListOptions Load event
	function ListOptions_Load() {

		// Example:
		$opt = &$this->ListOptions->Add("salenow");
		$opt->Header = "Sale Now";
		$opt->OnLeft = TRUE; // Link on left
		$opt->MoveTo(0); // Move to first column
	}

	// ListOptions Rendered event
	function ListOptions_Rendered() {

		// Example: 
		$this->ListOptions->Items["salenow"]->Body = "<a href='a_salesadd.php?Customer_Number=".$this->Customer_Number->CurrentValue."&showdetail=a_sales_detail'><span data-phrase='SaleLink' class='icon-new ewIcon' data-caption=''> Sale Now</span></a>";
	}

	// Row Custom Action event
	function Row_CustomAction() {

		// Return FALSE to abort
		return TRUE;
	}

	// Page Exporting event
	// $this->ExportDoc = export document object
	function Page_Exporting() {

		//$this->ExportDoc->Text = "my header"; // Export header
		//return FALSE; // Return FALSE to skip default export and use Row_Export event

		return TRUE; // Return TRUE to use default export and skip Row_Export event
	}

	// Row Export event
	// $this->ExportDoc = export document object
	function Row_Export() {

	    //$this->ExportDoc->Text .= "my content"; // Build HTML with field value: $rs["MyField"] or $this->MyField->ViewValue
	}

	// Page Exported event
	// $this->ExportDoc = export document object
	function Page_Exported() {

		//$this->ExportDoc->Text .= "my footer"; // Export footer
		//echo $this->ExportDoc->Text;

	}
}
?>
<?php ew_Header(FALSE) ?>
<?php

// Create page object
if (!isset($a_customers_list)) $a_customers_list = new ca_customers_list();

// Page init
$a_customers_list->Page_Init();

// Page main
$a_customers_list->Page_Main();

// Begin of modification Displaying Breadcrumb Links in All Pages, by Masino Sinaga, May 4, 2012
getCurrentPageTitle(ew_CurrentPage());

// End of modification Displaying Breadcrumb Links in All Pages, by Masino Sinaga, May 4, 2012
// Global Page Rendering event (in userfn*.php)

Page_Rendering();

// Global auto switch table width style (in userfn*.php), by Masino Sinaga, January 7, 2015
AutoSwitchTableWidthStyle();

// Page Rendering event
$a_customers_list->Page_Render();
?>
<?php include_once "header.php" ?>
<?php if ($a_customers->Export == "") { ?>
<script type="text/javascript">

// Form object
var CurrentPageID = EW_PAGE_ID = "list";
var CurrentForm = fa_customerslist = new ew_Form("fa_customerslist", "list");
fa_customerslist.FormKeyCountName = '<?php echo $a_customers_list->FormKeyCountName ?>';

// Form_CustomValidate event
fa_customerslist.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!

 	// Your custom validation code here, return false if invalid. 
 	return true;
 }

// Use JavaScript validation or not
<?php if (EW_CLIENT_VALIDATE) { ?>
fa_customerslist.ValidateRequired = true;
<?php } else { ?>
fa_customerslist.ValidateRequired = false; 
<?php } ?>

// Dynamic selection lists
// Form object for search

var CurrentSearchForm = fa_customerslistsrch = new ew_Form("fa_customerslistsrch");

// Init search panel as collapsed
<?php if (MS_USE_TABLE_SETTING_FOR_SEARCH_PANEL_COLLAPSED) { ?>
if (fa_customerslistsrch) fa_customerslistsrch.InitSearchPanel = true;
<?php } else { ?>
<?php if (MS_SEARCH_PANEL_COLLAPSED == TRUE && CurrentPage()->SearchWhere == "") { ?>
if (fa_customerslistsrch) fa_customerslistsrch.InitSearchPanel = true;
<?php } elseif ( (MS_SEARCH_PANEL_COLLAPSED == TRUE && CurrentPage()->SearchWhere <> "") || (MS_SEARCH_PANEL_COLLAPSED == FALSE && CurrentPage()->SearchWhere == "") ) { ?>
if (fa_customerslistsrch) fa_customerslistsrch.InitSearchPanel = false;
<?php } ?>
<?php } ?>
</script>
<script type="text/javascript" src="phpjs/ewscrolltable.min.js"></script>
<style type="text/css">
.ewTablePreviewRow { /* main table preview row color */
	background-color: #FFFFFF; /* preview row color */
}
.ewTablePreviewRow .ewGrid {
	display: table;
}
.ewTablePreviewRow .ewGrid .ewTable {
	width: auto;
}
</style>
<div id="ewPreview" class="hide"><ul class="nav nav-tabs"></ul><div class="tab-content"><div class="tab-pane fade"></div></div></div>
<script type="text/javascript" src="phpjs/ewpreview.min.js"></script>
<script type="text/javascript">
var EW_PREVIEW_PLACEMENT = EW_CSS_FLIP ? "left" : "right";
var EW_PREVIEW_SINGLE_ROW = false;
var EW_PREVIEW_OVERLAY = true;
</script>
<script type="text/javascript">

// Write your client script here, no need to add script tags.
</script>
<?php } ?>
<?php if ($a_customers->Export == "") { ?>
<?php $bShowLangSelector = false; ?>
<div class="ewToolbar">
<?php if ($a_customers->Export == "") { ?>
<?php if (MS_SHOW_PHPMAKER_BREADCRUMBLINKS) { ?>
<?php $Breadcrumb->Render(); ?>
<?php } ?>
<?php if (MS_SHOW_MASINO_BREADCRUMBLINKS) { ?>
<?php echo htmlspecialchars(MasinoBreadcrumbLinks()); ?>
<?php } ?>
<?php } ?>
<?php if ($a_customers_list->TotalRecs > 0 && $a_customers_list->ExportOptions->Visible()) { ?>
<?php $a_customers_list->ExportOptions->RenderOptions("body") ?>
<?php } ?>
<?php if ($bShowLangSelector == false) { ?>
<?php if ($a_customers_list->SearchOptions->Visible()) { ?>
<?php $a_customers_list->SearchOptions->RenderOptions("body") ?>
<?php } ?>
<?php if ($a_customers_list->FilterOptions->Visible()) { ?>
<?php $a_customers_list->FilterOptions->RenderOptions("body") ?>
<?php } ?>
<?php if ($a_customers->Export == "") { ?>
<?php if (MS_LANGUAGE_SELECTOR_VISIBILITY=="belowheader") { ?>
<?php echo $_SESSION['Language']->SelectionForm(); ?>
<?php } ?>
<?php } ?>
<?php } ?>
<div class="clearfix"></div>
</div>
<?php } ?>
<?php // movedown htmmaster session to htmheader session in template ?>
<?php
	$bSelectLimit = $a_customers_list->UseSelectLimit;
	if ($bSelectLimit) { // begin of v11.0.4
		if ($a_customers_list->TotalRecs <= 0)
			$a_customers_list->TotalRecs = $a_customers->SelectRecordCount();
	} else {
		if (!$a_customers_list->Recordset && ($a_customers_list->Recordset = $a_customers_list->LoadRecordset()))
			$a_customers_list->TotalRecs = $a_customers_list->Recordset->RecordCount();
	} // end of v11.0.4
	$a_customers_list->StartRec = 1;

// Begin of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012     
    if ($a_customers_list->DisplayRecs <= 0 || ($a_customers->Export <> "" && $a_customers->ExportAll=="allpages")) // Display all records
        $a_customers_list->DisplayRecs = $a_customers_list->TotalRecs;
    if (!($a_customers->Export <> "" && $a_customers->ExportAll=="allpages"))
        $a_customers_list->SetUpStartRec(); // Set up start record position

// End of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012
	if ($bSelectLimit)
		$a_customers_list->Recordset = $a_customers_list->LoadRecordset();

	// Set no record found message
	if ($a_customers->CurrentAction == "" && $a_customers_list->TotalRecs == 0) {
		if (!$_SESSION['Security']->CanList())
			$a_customers_list->setWarningMessage($_SESSION['Language']->Phrase("NoPermission"));
		if ($a_customers_list->SearchWhere == "0=101")
			$a_customers_list->setWarningMessage($_SESSION['Language']->Phrase("EnterSearchCriteria"));
		else
			$a_customers_list->setWarningMessage($_SESSION['Language']->Phrase("NoRecord"));
	}
$a_customers_list->RenderOtherOptions();
?>
<?php if ($_SESSION['Security']->CanSearch()) { ?>
<?php if ($a_customers->Export == "" && $a_customers->CurrentAction == "") { ?>
<form name="fa_customerslistsrch" id="fa_customerslistsrch" class="form-inline ewForm" action="<?php echo ew_CurrentPage() ?>">
<?php $SearchPanelClass = ($a_customers_list->SearchWhere <> "") ? " in" : " in"; ?>
<div id="fa_customerslistsrch_SearchPanel" class="ewSearchPanel collapse<?php echo $SearchPanelClass ?>">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="a_customers">
	<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="ewQuickSearch input-group">
	<input type="text" name="<?php echo EW_TABLE_BASIC_SEARCH ?>" id="<?php echo EW_TABLE_BASIC_SEARCH ?>" class="form-control" value="<?php echo ew_HtmlEncode($a_customers_list->BasicSearch->getKeyword()) ?>" placeholder="<?php echo ew_HtmlEncode($_SESSION['Language']->Phrase("Search")) ?>">
	<input type="hidden" name="<?php echo EW_TABLE_BASIC_SEARCH_TYPE ?>" id="<?php echo EW_TABLE_BASIC_SEARCH_TYPE ?>" value="<?php echo ew_HtmlEncode($a_customers_list->BasicSearch->getType()) ?>">
	<div class="input-group-btn">
		<button type="button" data-toggle="dropdown" class="btn btn-default"><span id="searchtype"><?php echo $a_customers_list->BasicSearch->getTypeNameShort() ?></span><span class="caret"></span></button>
		<ul class="dropdown-menu pull-right" role="menu">
			<li<?php if ($a_customers_list->BasicSearch->getType() == "") echo " class=\"active\""; ?>><a href="javascript:void(0);" onclick="ew_SetSearchType(this)"><?php echo $_SESSION['Language']->Phrase("QuickSearchAuto") ?></a></li>
			<li<?php if ($a_customers_list->BasicSearch->getType() == "=") echo " class=\"active\""; ?>><a href="javascript:void(0);" onclick="ew_SetSearchType(this,'=')"><?php echo $_SESSION['Language']->Phrase("QuickSearchExact") ?></a></li>
			<li<?php if ($a_customers_list->BasicSearch->getType() == "AND") echo " class=\"active\""; ?>><a href="javascript:void(0);" onclick="ew_SetSearchType(this,'AND')"><?php echo $_SESSION['Language']->Phrase("QuickSearchAll") ?></a></li>
			<li<?php if ($a_customers_list->BasicSearch->getType() == "OR") echo " class=\"active\""; ?>><a href="javascript:void(0);" onclick="ew_SetSearchType(this,'OR')"><?php echo $_SESSION['Language']->Phrase("QuickSearchAny") ?></a></li>
		</ul>
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><?php echo $_SESSION['Language']->Phrase("QuickSearchBtn") ?></button>
	</div>
	</div>
</div>
	</div>
</div>
</form>
<?php } ?>
<?php } ?>
<?php $a_customers_list->ShowPageHeader(); ?>
<?php
$a_customers_list->ShowMessage();
?>
<?php //////////////////////////// BEGIN Empty Table ?>
<?php // Begin of modification Displaying Empty Table, by Masino Sinaga, May 3, 2012 ?>
<?php if (MS_SHOW_EMPTY_TABLE_ON_LIST_PAGE) { ?>
<?php if ($a_customers_list->TotalRecs == 0) { ?>
<div class="panel panel-default ewGrid">
<?php if (MS_PAGINATION_POSITION == 1 || MS_PAGINATION_POSITION == 3) { ?>
<div class="panel-heading ewGridUpperPanel" style="height: 40px;">
<?php if ($a_customers_list->TotalRecs == 0 && $a_customers->CurrentAction == "") { // Show other options ?>
<div class="ewListOtherOptions">
<?php
	foreach ($a_customers_list->OtherOptions as &$option) {
		$option->ButtonClass = "";
		$option->RenderOptions("body", "");
	}
?>
</div>
<div class="clearfix"></div>
<?php } ?>
<div class="clearfix"></div><div class="ewPager"></div>
</div>
<?php } ?>
<div id="gmp_a_customers_empty_table" class="<?php if (ew_IsResponsiveLayout()) { echo "table-responsive "; } ?>ewGridMiddlePanel">
<table id="tbl_a_customerslist" class="table ewTable">
<?php echo $a_customers->TableCustomInnerHtml ?>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
    <?php
        $html6= new cod_html_6();
        echo $html6->html06($a_customers, $_SESSION['Language']);
    ?>
	</tr>
</thead>
<tbody>
	<tr<?php echo $a_customers->RowAttributes() ?>>
	    <?php
        $html5= new cod_html_5();
        echo $html5->html5($a_customers,$a_customers_list);
        ?>
	</tr>
</tbody>
<tfoot><!-- Table footer -->
	<tr class="ewTableFooter">
	<?php if ($a_customers->Customer_Number->Visible) { // Customer_Number ?>
		<td data-name="Customer_Number"><span id="elf_a_customers_Customer_Number" class="a_customers_Customer_Number">
		&nbsp;
		</span></td>
	<?php } ?>
	<?php if ($a_customers->Customer_Name->Visible) { // Customer_Name ?>
		<td data-name="Customer_Name"><span id="elf_a_customers_Customer_Name" class="a_customers_Customer_Name">
		&nbsp;
		</span></td>
	<?php } ?>
	<?php if ($a_customers->Contact_Person->Visible) { // Contact_Person ?>
		<td data-name="Contact_Person"><span id="elf_a_customers_Contact_Person" class="a_customers_Contact_Person">
		&nbsp;
		</span></td>
	<?php } ?>
	<?php if ($a_customers->Phone_Number->Visible) { // Phone_Number ?>
		<td data-name="Phone_Number"><span id="elf_a_customers_Phone_Number" class="a_customers_Phone_Number">
		&nbsp;
		</span></td>
	<?php } ?>
	<?php if ($a_customers->Mobile_Number->Visible) { // Mobile_Number ?>
		<td data-name="Mobile_Number"><span id="elf_a_customers_Mobile_Number" class="a_customers_Mobile_Number">
		&nbsp;
		</span></td>
	<?php } ?>
	<?php if ($a_customers->Balance->Visible) { // Balance ?>
		<td data-name="Balance"><span id="elf_a_customers_Balance" class="a_customers_Balance">
<span class="ewAggregate"><?php echo $_SESSION['Language']->Phrase("TOTAL") ?></span>
<?php echo $a_customers->Balance->ViewValue ?>
		</span></td>
	<?php } ?>
	</tr>
</tfoot>
</table>
</div>
<?php if (MS_PAGINATION_POSITION == 2 || MS_PAGINATION_POSITION == 3) { ?>
<div class="panel-footer ewGridLowerPanel" style="height: 40px;">
<?php if ($a_customers_list->TotalRecs == 0 && $a_customers->CurrentAction == "") { // Show other options ?>
<div class="ewListOtherOptions">
<?php
	foreach ($a_customers_list->OtherOptions as &$option) {
		$option->ButtonClass = "";
		$option->RenderOptions("body", "");
	}
?>
</div>
<div class="clearfix"></div>
<?php } ?>
<div class="clearfix"></div></div>
<?php } ?>
</div>
<?php } ?>
<?php } ?>
<?php // End of modification Displaying Empty Table, by Masino Sinaga, May 3, 2012 ?>
<?php //////////////////////////// END Empty Table ?>
<?php if ($a_customers_list->TotalRecs > 0 || $a_customers->CurrentAction <> "") { ?>
<div class="panel panel-default ewGrid">
<?php // Begin of modification Customize Navigation/Pager Panel, by Masino Sinaga, May 2, 2012 ?>
<?php if ( (MS_PAGINATION_POSITION==1) || (MS_PAGINATION_POSITION==3) ) { ?>
<?php if ($a_customers->Export == "") { ?>
<div class="panel-heading ewGridUpperPanel">
<?php if ($a_customers->CurrentAction <> "gridadd" && $a_customers->CurrentAction <> "gridedit") { ?>
<form name="ewPagerForm" class="form-inline ewForm ewPagerForm" action="<?php echo ew_CurrentPage() ?>">
<?php if ($a_customers_list->TotalRecs > 0 && ((MS_SELECTABLE_PAGE_SIZES_POSITION=="Left" && $_SESSION['Language']->Phrase("dir")!="rtl") || (MS_SELECTABLE_PAGE_SIZES_POSITION=="Left" && $_SESSION['Language']->Phrase("dir")=="rtl"))) { ?>
<div class="ewPager"><span>&nbsp;<?php echo $_SESSION['Language']->Phrase("RecordsPerPage") ?>&nbsp;</span>
<input type="hidden" name="t" value="a_customers">
<select name="<?php echo EW_TABLE_REC_PER_PAGE ?>" class="form-control input-sm" onchange="this.form.submit();">
<?php $sRecPerPageList = explode(',', MS_TABLE_SELECTABLE_REC_PER_PAGE_LIST); ?>
<?php
foreach ($sRecPerPageList as $a) {
    $thisDisplayRecs = $a;
    ?>
    <option value="<?php if($thisDisplayRecs > 0 ){echo $thisDisplayRecs;}else{echo "ALL";}?>" <?php if(($thisDisplayRecs > 0 && $a_customers_list->DisplayRecs == $thisDisplayRecs) || $a_customers->getRecordsPerPage() == -1){ ?>selected="selected"<?php } ?>><?php if($thisDisplayRecs > 0){echo $thisDisplayRecs;}else{echo $_SESSION['Language']->Phrase("AllRecords");}?></option>
<?php } ?>
</select>
</div>
<?php } ?>
		<?php if (MS_PAGINATION_STYLE==1 && !isset($a_customers_list->Pager)) $a_customers_list->Pager = new cNumericPager($a_customers_list->StartRec, $a_customers_list->DisplayRecs, $a_customers_list->TotalRecs, $a_customers_list->RecRange) ?>
				<?php if ((MS_PAGINATION_STYLE==1 && $a_customers_list->Pager->RecordCount > 0) && (($a_customers_list->Pager->PageCount!=1) || ($a_customers_list->Pager->CurrentPage != 1) || (MS_SHOW_PAGENUM_IF_REC_NOT_OVER_PAGESIZE==TRUE)) ) { ?>
				    <?php
                    $html= new cod_html();
				    echo $html->codhtml($a_customers_list, $_SESSION['Language']);
				    ?>
				<?php } // end MS_SHOW_PAGENUM_IF_REC_NOT_OVER_PAGESIZE ?>

        <?php if (MS_PAGINATION_STYLE==1 && $a_customers_list->Pager->RecordCount > 0) { ?>
            <?php
                $html2= new cod_html();
                echo $html2->codhtml2($_SESSION['Language'], $a_customers_list);
            ?>
		<?php } ?>

		<?php if (MS_PAGINATION_STYLE==2 && !isset($a_customers_list->Pager)) $a_customers_list->Pager = new cPrevNextPager($a_customers_list->StartRec, $a_customers_list->DisplayRecs, $a_customers_list->TotalRecs) ?>
				<?php
                $html=new cod_html();
                $html->codhtml($a_customers_list,$_SESSION['Language']);
                ?>
        <?php if (MS_PAGINATION_STYLE==2 && $a_customers_list->Pager->RecordCount > 0) { ?>
            <?php
            $html2=new cod_html();
            $html2->codhtml2($a_customers_list,$_SESSION['Language']);
            ?>
		<?php } ?>

<?php if ( ($a_customers_list->TotalRecs > 0 && (MS_SELECTABLE_PAGE_SIZES_POSITION=="Right" && $_SESSION['Language']->Phrase("dir")!="rtl") || (MS_SELECTABLE_PAGE_SIZES_POSITION=="Right" && $_SESSION['Language']->Phrase("dir")=="rtl"))) { ?>
<div class="ewPager"><span>&nbsp;<?php echo $_SESSION['Language']->Phrase("RecordsPerPage") ?>&nbsp;</span>
<input type="hidden" name="t" value="a_customers">
<select name="<?php echo EW_TABLE_REC_PER_PAGE ?>" class="form-control input-sm" onchange="this.form.submit();">
<option value="1"<?php if ($a_customers_list->DisplayRecs == 1) { ?> selected="selected"<?php } ?>>1</option>
<option value="3"<?php if ($a_customers_list->DisplayRecs == 3) { ?> selected="selected"<?php } ?>>3</option>
<option value="5"<?php if ($a_customers_list->DisplayRecs == 5) { ?> selected="selected"<?php } ?>>5</option>
<option value="10"<?php if ($a_customers_list->DisplayRecs == 10) { ?> selected="selected"<?php } ?>>10</option>
<option value="20"<?php if ($a_customers_list->DisplayRecs == 20) { ?> selected="selected"<?php } ?>>20</option>
<option value="50"<?php if ($a_customers_list->DisplayRecs == 50) { ?> selected="selected"<?php } ?>>50</option>
<option value="100"<?php if ($a_customers_list->DisplayRecs == 100) { ?> selected="selected"<?php } ?>>100</option>
</select>
</div>
<?php } // end if (MS_SELECTABLE_PAGE_SIZES_POSITION=="Right") ?>

</form>
<?php } ?>
<div class="ewListOtherOptions">
<?php
	foreach ($a_customers_list->OtherOptions as &$option)
		$option->RenderOptions("body");
?>
</div>
<div class="clearfix"></div>
</div>
<?php } ?>
<?php } ?>
<?php // End of modification Customize Navigation/Pager Panel, by Masino Sinaga, May 2, 2012 ?>
<form name="fa_customerslist" id="fa_customerslist" class="form-inline ewForm ewListForm" action="<?php echo ew_CurrentPage() ?>" method="post">
<?php if ($a_customers_list->CheckToken) { ?>
<input type="hidden" name="<?php echo EW_TOKEN_NAME ?>" value="<?php echo $a_customers_list->Token ?>">
<?php } ?>
<input type="hidden" name="t" value="a_customers">
<?php // Begin of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012 ?>
<?php if (MS_EXPORT_RECORD_OPTIONS=="selectedrecords") { ?>
<input type="hidden" name="exporttype" id="exporttype" value="">
<?php } ?>
<?php // End of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012 ?>
<div id="gmp_a_customers" class="<?php if (ew_IsResponsiveLayout()) { echo "table-responsive "; } ?>ewGridMiddlePanel">
<?php if ($a_customers_list->TotalRecs > 0) { ?>
<table id="tbl_a_customerslist" class="table ewTable">
<?php echo $a_customers->TableCustomInnerHtml ?>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<?php

// Header row
$a_customers_list->RowType = EW_ROWTYPE_HEADER; // since v11.0.6

// Render list options
$a_customers_list->RenderListOptions();

// Render list options (header, left)
$a_customers_list->ListOptions->RenderOptions("header", "left");
?>
<?php if ($a_customers->Customer_Number->Visible) { // Customer_Number ?>
	<?php if ($a_customers->SortUrl($a_customers->Customer_Number) == "") { ?>
		<th data-name="Customer_Number"><div id="elh_a_customers_Customer_Number" class="a_customers_Customer_Number"><div class="ewTableHeaderCaption"><?php echo $a_customers->Customer_Number->FldCaption() ?></div></div></th>
	<?php } else { ?>
		<th data-name="Customer_Number"><div class="ewPointer" onclick="ew_Sort(event,'<?php echo $a_customers->SortUrl($a_customers->Customer_Number) ?>',1);"><div id="elh_a_customers_Customer_Number" class="a_customers_Customer_Number">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><?php echo $a_customers->Customer_Number->FldCaption() ?><?php echo $_SESSION['Language']->Phrase("SrchLegend") ?></span><span class="ewTableHeaderSort"><?php if ($a_customers->Customer_Number->getSort() == "ASC") { ?><span class="caret ewSortUp"></span><?php } elseif ($a_customers->Customer_Number->getSort() == "DESC") { ?><span class="caret"></span><?php } ?></span></div>
        </div></div></th>
	<?php } ?>
<?php } ?>		
<?php if ($a_customers->Customer_Name->Visible) { // Customer_Name ?>
	<?php if ($a_customers->SortUrl($a_customers->Customer_Name) == "") { ?>
		<th data-name="Customer_Name"><div id="elh_a_customers_Customer_Name" class="a_customers_Customer_Name"><div class="ewTableHeaderCaption"><?php echo $a_customers->Customer_Name->FldCaption() ?></div></div></th>
	<?php } else { ?>
		<th data-name="Customer_Name"><div class="ewPointer" onclick="ew_Sort(event,'<?php echo $a_customers->SortUrl($a_customers->Customer_Name) ?>',1);"><div id="elh_a_customers_Customer_Name" class="a_customers_Customer_Name">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><?php echo $a_customers->Customer_Name->FldCaption() ?><?php echo $_SESSION['Language']->Phrase("SrchLegend") ?></span><span class="ewTableHeaderSort"><?php if ($a_customers->Customer_Name->getSort() == "ASC") { ?><span class="caret ewSortUp"></span><?php } elseif ($a_customers->Customer_Name->getSort() == "DESC") { ?><span class="caret"></span><?php } ?></span></div>
        </div></div></th>
	<?php } ?>
<?php } ?>		
<?php if ($a_customers->Contact_Person->Visible) { // Contact_Person ?>
	<?php if ($a_customers->SortUrl($a_customers->Contact_Person) == "") { ?>
		<th data-name="Contact_Person"><div id="elh_a_customers_Contact_Person" class="a_customers_Contact_Person"><div class="ewTableHeaderCaption"><?php echo $a_customers->Contact_Person->FldCaption() ?></div></div></th>
	<?php } else { ?>
		<th data-name="Contact_Person"><div class="ewPointer" onclick="ew_Sort(event,'<?php echo $a_customers->SortUrl($a_customers->Contact_Person) ?>',1);"><div id="elh_a_customers_Contact_Person" class="a_customers_Contact_Person">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><?php echo $a_customers->Contact_Person->FldCaption() ?><?php echo $_SESSION['Language']->Phrase("SrchLegend") ?></span><span class="ewTableHeaderSort"><?php if ($a_customers->Contact_Person->getSort() == "ASC") { ?><span class="caret ewSortUp"></span><?php } elseif ($a_customers->Contact_Person->getSort() == "DESC") { ?><span class="caret"></span><?php } ?></span></div>
        </div></div></th>
	<?php } ?>
<?php } ?>		
<?php if ($a_customers->Phone_Number->Visible) { // Phone_Number ?>
	<?php if ($a_customers->SortUrl($a_customers->Phone_Number) == "") { ?>
		<th data-name="Phone_Number"><div id="elh_a_customers_Phone_Number" class="a_customers_Phone_Number"><div class="ewTableHeaderCaption"><?php echo $a_customers->Phone_Number->FldCaption() ?></div></div></th>
	<?php } else { ?>
		<th data-name="Phone_Number"><div class="ewPointer" onclick="ew_Sort(event,'<?php echo $a_customers->SortUrl($a_customers->Phone_Number) ?>',1);"><div id="elh_a_customers_Phone_Number" class="a_customers_Phone_Number">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><?php echo $a_customers->Phone_Number->FldCaption() ?><?php echo $_SESSION['Language']->Phrase("SrchLegend") ?></span><span class="ewTableHeaderSort"><?php if ($a_customers->Phone_Number->getSort() == "ASC") { ?><span class="caret ewSortUp"></span><?php } elseif ($a_customers->Phone_Number->getSort() == "DESC") { ?><span class="caret"></span><?php } ?></span></div>
        </div></div></th>
	<?php } ?>
<?php } ?>		
<?php if ($a_customers->Mobile_Number->Visible) { // Mobile_Number ?>
	<?php if ($a_customers->SortUrl($a_customers->Mobile_Number) == "") { ?>
		<th data-name="Mobile_Number"><div id="elh_a_customers_Mobile_Number" class="a_customers_Mobile_Number"><div class="ewTableHeaderCaption"><?php echo $a_customers->Mobile_Number->FldCaption() ?></div></div></th>
	<?php } else { ?>
		<th data-name="Mobile_Number"><div class="ewPointer" onclick="ew_Sort(event,'<?php echo $a_customers->SortUrl($a_customers->Mobile_Number) ?>',1);"><div id="elh_a_customers_Mobile_Number" class="a_customers_Mobile_Number">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><?php echo $a_customers->Mobile_Number->FldCaption() ?><?php echo $_SESSION['Language']->Phrase("SrchLegend") ?></span><span class="ewTableHeaderSort"><?php if ($a_customers->Mobile_Number->getSort() == "ASC") { ?><span class="caret ewSortUp"></span><?php } elseif ($a_customers->Mobile_Number->getSort() == "DESC") { ?><span class="caret"></span><?php } ?></span></div>
        </div></div></th>
	<?php } ?>
<?php } ?>		
<?php if ($a_customers->Balance->Visible) { // Balance ?>
	<?php if ($a_customers->SortUrl($a_customers->Balance) == "") { ?>
		<th data-name="Balance"><div id="elh_a_customers_Balance" class="a_customers_Balance"><div class="ewTableHeaderCaption"><?php echo $a_customers->Balance->FldCaption() ?></div></div></th>
	<?php } else { ?>
		<th data-name="Balance"><div class="ewPointer" onclick="ew_Sort(event,'<?php echo $a_customers->SortUrl($a_customers->Balance) ?>',1);"><div id="elh_a_customers_Balance" class="a_customers_Balance">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><?php echo $a_customers->Balance->FldCaption() ?><?php echo $_SESSION['Language']->Phrase("SrchLegend") ?></span><span class="ewTableHeaderSort"><?php if ($a_customers->Balance->getSort() == "ASC") { ?><span class="caret ewSortUp"></span><?php } elseif ($a_customers->Balance->getSort() == "DESC") { ?><span class="caret"></span><?php } ?></span></div>
        </div></div></th>
	<?php } ?>
<?php } ?>		
<?php

// Render list options (header, right)
$a_customers_list->ListOptions->RenderOptions("header", "right");
?>
	</tr>
</thead>
<tbody>
<?php

// Begin of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012
if ($a_customers->ExportAll=="allpages" && $a_customers->Export <> "") {
    $a_customers_list->StopRec = $a_customers_list->TotalRecs;

// End of mofidication Flexibility of Export Records Options, by Masino Sinaga, May 14, 2012
} else {

	// Set the last record to display
	if ($a_customers_list->TotalRecs > $a_customers_list->StartRec + $a_customers_list->DisplayRecs - 1)
		$a_customers_list->StopRec = $a_customers_list->StartRec + $a_customers_list->DisplayRecs - 1;
	else
		$a_customers_list->StopRec = $a_customers_list->TotalRecs;
}
$a_customers_list->RecCnt = $a_customers_list->StartRec - 1;
if ($a_customers_list->Recordset && !$a_customers_list->Recordset->EOF) {
	$a_customers_list->Recordset->MoveFirst();
	$bSelectLimit = $a_customers_list->UseSelectLimit;
	if (!$bSelectLimit && $a_customers_list->StartRec > 1)
		$a_customers_list->Recordset->Move($a_customers_list->StartRec - 1);
} elseif (!$a_customers->AllowAddDeleteRow && $a_customers_list->StopRec == 0) {
	$a_customers_list->StopRec = $a_customers->GridAddRowCount;
}

// Initialize aggregate
$a_customers->RowType = EW_ROWTYPE_AGGREGATEINIT;
$a_customers->ResetAttrs();
$a_customers_list->RenderRow();
while ($a_customers_list->RecCnt < $a_customers_list->StopRec) {
	$a_customers_list->RecCnt++;
	if (intval($a_customers_list->RecCnt) >= intval($a_customers_list->StartRec)) {
		$a_customers_list->RowCnt++;

		// Set up key count
		$a_customers_list->KeyCount = $a_customers_list->RowIndex;

		// Init row class and style
		$a_customers->ResetAttrs();
		$a_customers->CssClass = "";
		if ($a_customers->CurrentAction == "gridadd") {
		} else {
			$a_customers_list->LoadRowValues($a_customers_list->Recordset); // Load row values
		}
		$a_customers->RowType = EW_ROWTYPE_VIEW; // Render view

		// Set up row id / data-rowindex
		$a_customers->RowAttrs = array_merge($a_customers->RowAttrs, array('data-rowindex'=>$a_customers_list->RowCnt, 'id'=>'r' . $a_customers_list->RowCnt . '_a_customers', 'data-rowtype'=>$a_customers->RowType));

		// Render row
		$a_customers_list->RenderRow();

		// Render list options
		$a_customers_list->RenderListOptions();
?>
	<tr<?php echo $a_customers->RowAttributes() ?>>
<?php

// Render list options (body, left)
$a_customers_list->ListOptions->RenderOptions("body", "left", $a_customers_list->RowCnt);
?>
	<?php if ($a_customers->Customer_Number->Visible) { // Customer_Number ?>
		<td data-name="Customer_Number"<?php echo $a_customers->Customer_Number->CellAttributes() ?>>
<span id="el<?php echo $a_customers_list->RowCnt ?>_a_customers_Customer_Number" class="a_customers_Customer_Number">
<span<?php echo $a_customers->Customer_Number->ViewAttributes() ?>>
<?php echo $a_customers->Customer_Number->ListViewValue() ?></span>
</span>
<a id="<?php echo $a_customers_list->PageObjName . "_row_" . $a_customers_list->RowCnt ?>"></a></td>
	<?php } ?>
	<?php if ($a_customers->Customer_Name->Visible) { // Customer_Name ?>
		<td data-name="Customer_Name"<?php echo $a_customers->Customer_Name->CellAttributes() ?>>
<span id="el<?php echo $a_customers_list->RowCnt ?>_a_customers_Customer_Name" class="a_customers_Customer_Name">
<span<?php echo $a_customers->Customer_Name->ViewAttributes() ?>>
<?php echo $a_customers->Customer_Name->ListViewValue() ?></span>
</span>
</td>
	<?php } ?>
	<?php if ($a_customers->Contact_Person->Visible) { // Contact_Person ?>
		<td data-name="Contact_Person"<?php echo $a_customers->Contact_Person->CellAttributes() ?>>
<span id="el<?php echo $a_customers_list->RowCnt ?>_a_customers_Contact_Person" class="a_customers_Contact_Person">
<span<?php echo $a_customers->Contact_Person->ViewAttributes() ?>>
<?php echo $a_customers->Contact_Person->ListViewValue() ?></span>
</span>
</td>
	<?php } ?>
	<?php if ($a_customers->Phone_Number->Visible) { // Phone_Number ?>
		<td data-name="Phone_Number"<?php echo $a_customers->Phone_Number->CellAttributes() ?>>
<span id="el<?php echo $a_customers_list->RowCnt ?>_a_customers_Phone_Number" class="a_customers_Phone_Number">
<span<?php echo $a_customers->Phone_Number->ViewAttributes() ?>>
<?php echo $a_customers->Phone_Number->ListViewValue() ?></span>
</span>
</td>
	<?php } ?>
	<?php if ($a_customers->Mobile_Number->Visible) { // Mobile_Number ?>
		<td data-name="Mobile_Number"<?php echo $a_customers->Mobile_Number->CellAttributes() ?>>
<span id="el<?php echo $a_customers_list->RowCnt ?>_a_customers_Mobile_Number" class="a_customers_Mobile_Number">
<span<?php echo $a_customers->Mobile_Number->ViewAttributes() ?>>
<?php echo $a_customers->Mobile_Number->ListViewValue() ?></span>
</span>
</td>
	<?php } ?>
	<?php if ($a_customers->Balance->Visible) { // Balance ?>
		<td data-name="Balance"<?php echo $a_customers->Balance->CellAttributes() ?>>
<span id="el<?php echo $a_customers_list->RowCnt ?>_a_customers_Balance" class="a_customers_Balance">
<span<?php echo $a_customers->Balance->ViewAttributes() ?>>
<?php echo $a_customers->Balance->ListViewValue() ?></span>
</span>
</td>
	<?php } ?>
<?php

// Render list options (body, right)
$a_customers_list->ListOptions->RenderOptions("body", "right", $a_customers_list->RowCnt);
?>
	</tr>
<?php
	}
	if ($a_customers->CurrentAction <> "gridadd")
		$a_customers_list->Recordset->MoveNext();
}
?>
</tbody>
<?php

// Render aggregate row
$a_customers->RowType = EW_ROWTYPE_AGGREGATE;
$a_customers->ResetAttrs();
$a_customers_list->RenderRow();
?>
<?php if ($a_customers_list->TotalRecs > 0 && ($a_customers->CurrentAction <> "gridadd" && $a_customers->CurrentAction <> "gridedit")) { ?>
<tfoot><!-- Table footer -->
	<tr class="ewTableFooter">
<?php

// Render list options
$a_customers_list->RenderListOptions();

// Render list options (footer, left)
$a_customers_list->ListOptions->RenderOptions("footer", "left");
?>
	<?php if ($a_customers->Customer_Number->Visible) { // Customer_Number ?>
		<td data-name="Customer_Number"><span id="elf_a_customers_Customer_Number" class="a_customers_Customer_Number">
		&nbsp;
		</span></td>
	<?php } ?>
	<?php if ($a_customers->Customer_Name->Visible) { // Customer_Name ?>
		<td data-name="Customer_Name"><span id="elf_a_customers_Customer_Name" class="a_customers_Customer_Name">
		&nbsp;
		</span></td>
	<?php } ?>
	<?php if ($a_customers->Contact_Person->Visible) { // Contact_Person ?>
		<td data-name="Contact_Person"><span id="elf_a_customers_Contact_Person" class="a_customers_Contact_Person">
		&nbsp;
		</span></td>
	<?php } ?>
	<?php if ($a_customers->Phone_Number->Visible) { // Phone_Number ?>
		<td data-name="Phone_Number"><span id="elf_a_customers_Phone_Number" class="a_customers_Phone_Number">
		&nbsp;
		</span></td>
	<?php } ?>
	<?php if ($a_customers->Mobile_Number->Visible) { // Mobile_Number ?>
		<td data-name="Mobile_Number"><span id="elf_a_customers_Mobile_Number" class="a_customers_Mobile_Number">
		&nbsp;
		</span></td>
	<?php } ?>
	<?php if ($a_customers->Balance->Visible) { // Balance ?>
		<td data-name="Balance"><span id="elf_a_customers_Balance" class="a_customers_Balance">
<span class="ewAggregate"><?php echo $_SESSION['Language']->Phrase("TOTAL") ?></span>
<?php echo $a_customers->Balance->ViewValue ?>
		</span></td>
	<?php } ?>
<?php

// Render list options (footer, right)
$a_customers_list->ListOptions->RenderOptions("footer", "right");
?>
	</tr>
</tfoot>	
<?php } ?>
</table>
<?php } ?>
<?php if ($a_customers->CurrentAction == "") { ?>
<input type="hidden" name="a_list" id="a_list" value="">
<?php } ?>
</div>
</form>
<?php

// Close recordset
if ($a_customers_list->Recordset)
	$a_customers_list->Recordset->Close();
?>
<?php // Begin of modification Customize Navigation/Pager Panel, by Masino Sinaga, May 2, 2012 ?>
<?php if ( (MS_PAGINATION_POSITION==2) || (MS_PAGINATION_POSITION==3) ) { ?>
<?php if ($a_customers->Export == "") { ?>
<div class="panel-footer ewGridLowerPanel">
<?php if ($a_customers->CurrentAction <> "gridadd" && $a_customers->CurrentAction <> "gridedit") { ?>
<form name="ewPagerForm" class="ewForm form-inline ewPagerForm" action="<?php echo ew_CurrentPage() ?>">
<?php if ($a_customers_list->TotalRecs > 0 && ((MS_SELECTABLE_PAGE_SIZES_POSITION=="Left" && $_SESSION['Language']->Phrase("dir")!="rtl") || (MS_SELECTABLE_PAGE_SIZES_POSITION=="Left" && $_SESSION['Language']->Phrase("dir")=="rtl"))) { ?>
<div class="ewPager"><span>&nbsp;<?php echo $_SESSION['Language']->Phrase("RecordsPerPage") ?>&nbsp;</span>
<input type="hidden" name="t" value="a_customers">
<select name="<?php echo EW_TABLE_REC_PER_PAGE ?>" class="form-control input-sm" onchange="this.form.submit();">
<?php $sRecPerPageList = explode(',', MS_TABLE_SELECTABLE_REC_PER_PAGE_LIST); ?>
<?php
foreach ($sRecPerPageList as $a) {
    $thisDisplayRecs = $a;
    ?>
    <option value="<?php if($thisDisplayRecs > 0 ){echo $thisDisplayRecs;}else{echo "ALL";}?>" <?php if(($thisDisplayRecs > 0 && $a_customers_list->DisplayRecs == $thisDisplayRecs) || $a_customers->getRecordsPerPage() == -1){ ?>selected="selected"<?php } ?>><?php if($thisDisplayRecs > 0){echo $thisDisplayRecs;}else{echo $_SESSION['Language']->Phrase("AllRecords");}?></option>
<?php } ?>

</select>
</div>
<?php } ?>
		<?php if (MS_PAGINATION_STYLE==1 && !isset($a_customers_list->Pager)) $a_customers_list->Pager = new cNumericPager($a_customers_list->StartRec, $a_customers_list->DisplayRecs, $a_customers_list->TotalRecs, $a_customers_list->RecRange) ?>
				<?php if (MS_PAGINATION_STYLE==1 && $a_customers_list->Pager->RecordCount > 0 && (($a_customers_list->Pager->PageCount!=1) || ($a_customers_list->Pager->CurrentPage != 1) || (MS_SHOW_PAGENUM_IF_REC_NOT_OVER_PAGESIZE==TRUE))) { ?>
				<?php
                    $html = new cod_html();
                    echo $html->codhtml($a_customers_list, $_SESSION['Language']);
                    ?>
				<?php } // end MS_SHOW_PAGENUM_IF_REC_NOT_OVER_PAGESIZE ?>
        <?php if (MS_PAGINATION_STYLE==1 && $a_customers_list->Pager->RecordCount > 0) { ?>
            <?php
            $html2 = new cod_html();
            echo $html2->codhtml2($a_customers_list, $_SESSION['Language']);
            ?>
		<?php } ?>
		<?php if (MS_PAGINATION_STYLE==2 && !isset($a_customers_list->Pager)) $a_customers_list->Pager = new cPrevNextPager($a_customers_list->StartRec, $a_customers_list->DisplayRecs, $a_customers_list->TotalRecs) ?>
				<?php if (($a_customers_list->Pager->PageCount!=1) || ($a_customers_list->Pager->CurrentPage != 1) || (MS_SHOW_PAGENUM_IF_REC_NOT_OVER_PAGESIZE==TRUE)  ) { ?>
				<div class="ewPager">
				<span><?php echo $_SESSION['Language']->Phrase("Page") ?>&nbsp;</span>
				<div class="ewPrevNext"><div class="input-group">
				<div class="input-group-btn">
				<!--first page button-->
                    <?php if ($a_customers_list->Pager->FirstButton->Enabled && $_SESSION['Language']->Phrase("dir") == "rtl") { ?>
                        <a class="btn btn-default btn-sm" title="<?php echo $_SESSION['Language']->Phrase("PagerFirst") ?>" href="<?php echo $a_customers_list->PageUrl() ?>start=<?php echo $a_customers_list->Pager->FirstButton->Start ?>"><span class="icon-last ewIcon"></span></a>
                    <?php } ?>
                    <?php if ($a_customers_list->Pager->FirstButton->Enabled && $_SESSION['Language']->Phrase("dir") != "rtl") { ?>
                        <a class="btn btn-default btn-sm" title="<?php echo $_SESSION['Language']->Phrase("PagerFirst") ?>" href="<?php echo $a_customers_list->PageUrl() ?>start=<?php echo $a_customers_list->Pager->FirstButton->Start ?>"><span class="icon-first ewIcon"></span></a>
                    <?php } ?>
                <?php if ((!$a_customers_list->Pager->FirstButton->Enabled) && $_SESSION['Language']->Phrase("dir") == "rtl") { ?>
                    <a class="btn btn-default btn-sm disabled" title="<?php echo $_SESSION['Language']->Phrase("PagerFirst") ?>"><span class="icon-last ewIcon"></span></a>
                <?php } ?>
                <?php if ((!$a_customers_list->Pager->FirstButton->Enabled) && $_SESSION['Language']->Phrase("dir") != "rtl") { ?>
                    <a class="btn btn-default btn-sm disabled" title="<?php echo $_SESSION['Language']->Phrase("PagerFirst") ?>"><span class="icon-first ewIcon"></span></a>
                <?php } ?>

				<!--previous page button-->

                    <?php if ($a_customers_list->Pager->PrevButton->Enabled && $_SESSION['Language']->Phrase("dir") == "rtl") { ?>
                        <a class="btn btn-default btn-sm" title="<?php echo $_SESSION['Language']->Phrase("PagerPrevious") ?>" href="<?php echo $a_customers_list->PageUrl() ?>start=<?php echo $a_customers_list->Pager->PrevButton->Start ?>"><span class="icon-next ewIcon"></span></a>
                    <?php }  ?>
                    <?php if ($a_customers_list->Pager->PrevButton->Enabled && $_SESSION['Language']->Phrase("dir") != "rtl") { ?>
                        <a class="btn btn-default btn-sm" title="<?php echo $_SESSION['Language']->Phrase("PagerPrevious") ?>" href="<?php echo $a_customers_list->PageUrl() ?>start=<?php echo $a_customers_list->Pager->PrevButton->Start ?>"><span class="icon-prev ewIcon"></span></a>
                    <?php }  ?>
                    <?php if ((!$a_customers_list->Pager->PrevButton->Enabled) && $_SESSION['Language']->Phrase("dir") == "rtl") { ?>
                        <a class="btn btn-default btn-sm disabled" title="<?php echo $_SESSION['Language']->Phrase("PagerPrevious") ?>"><span class="icon-next ewIcon"></span></a>
                    <?php }  ?>
                    <?php if ((!$a_customers_list->Pager->PrevButton->Enabled) && $_SESSION['Language']->Phrase("dir") != "rtl") { ?>
                        <a class="btn btn-default btn-sm disabled" title="<?php echo $_SESSION['Language']->Phrase("PagerPrevious") ?>"><span class="icon-prev ewIcon"></span></a>
                    <?php }  ?>

				</div>
				<!--current page number-->
					<input class="form-control input-sm" type="text" name="<?php echo EW_TABLE_PAGE_NO ?>" value="<?php echo $a_customers_list->Pager->CurrentPage ?>">
				<div class="input-group-btn">
				<!--next page button-->

                    <?php if ($a_customers_list->Pager->NextButton->Enabled && $_SESSION['Language']->Phrase("dir") == "rtl") { ?>
                        <a class="btn btn-default btn-sm" title="<?php echo $_SESSION['Language']->Phrase("PagerNext") ?>" href="<?php echo $a_customers_list->PageUrl() ?>start=<?php echo $a_customers_list->Pager->NextButton->Start ?>"><span class="icon-prev ewIcon"></span></a>
                    <?php } ?>
                    <?php if ($a_customers_list->Pager->NextButton->Enabled && $_SESSION['Language']->Phrase("dir") != "rtl") { ?>
                        <a class="btn btn-default btn-sm" title="<?php echo $_SESSION['Language']->Phrase("PagerNext") ?>" href="<?php echo $a_customers_list->PageUrl() ?>start=<?php echo $a_customers_list->Pager->NextButton->Start ?>"><span class="icon-next ewIcon"></span></a>
                    <?php } ?>
                    <?php if ((!$a_customers_list->Pager->NextButton->Enabled) && $_SESSION['Language']->Phrase("dir") == "rtl") { ?>
                        <a class="btn btn-default btn-sm disabled" title="<?php echo $_SESSION['Language']->Phrase("PagerNext") ?>"><span class="icon-prev ewIcon"></span></a>
                    <?php } ?>
                    <?php if ((!$a_customers_list->Pager->NextButton->Enabled) && $_SESSION['Language']->Phrase("dir") != "rtl") { ?>
                        <a class="btn btn-default btn-sm disabled" title="<?php echo $_SESSION['Language']->Phrase("PagerNext") ?>"><span class="icon-next ewIcon"></span></a>
                    <?php } ?>

				<!--last page button-->

                    <?php if ($a_customers_list->Pager->LastButton->Enabled && $_SESSION['Language']->Phrase("dir") == "rtl") { ?>
                        <a class="btn btn-default btn-sm" title="<?php echo $_SESSION['Language']->Phrase("PagerLast") ?>" href="<?php echo $a_customers_list->PageUrl() ?>start=<?php echo $a_customers_list->Pager->LastButton->Start ?>"><span class="icon-first ewIcon"></span></a>
                    <?php } ?>
                    <?php if ($a_customers_list->Pager->LastButton->Enabled && $_SESSION['Language']->Phrase("dir") != "rtl") { ?>
                        <a class="btn btn-default btn-sm" title="<?php echo $_SESSION['Language']->Phrase("PagerLast") ?>" href="<?php echo $a_customers_list->PageUrl() ?>start=<?php echo $a_customers_list->Pager->LastButton->Start ?>"><span class="icon-last ewIcon"></span></a>
                    <?php } ?>
                    <?php if ((!$a_customers_list->Pager->LastButton->Enabled) && $_SESSION['Language']->Phrase("dir") == "rtl") { ?>
                        <a class="btn btn-default btn-sm disabled" title="<?php echo $_SESSION['Language']->Phrase("PagerLast") ?>"><span class="icon-first ewIcon"></span></a>
                    <?php } ?>
                    <?php if ((!$a_customers_list->Pager->LastButton->Enabled) && $_SESSION['Language']->Phrase("dir") != "rtl") { ?>
                        <a class="btn btn-default btn-sm disabled" title="<?php echo $_SESSION['Language']->Phrase("PagerLast") ?>"><span class="icon-last ewIcon"></span></a>
                    <?php } ?>

				</div>
				</div>
				</div>
				<span>&nbsp;<?php echo $_SESSION['Language']->Phrase("of") ?>&nbsp;<?php echo $a_customers_list->Pager->PageCount ?></span>
				</div>
				<?php } // end MS_SHOW_PAGENUM_IF_REC_NOT_OVER_PAGESIZE==FALSE ?>
        <?php if (MS_PAGINATION_STYLE==2 && $a_customers_list->Pager->RecordCount > 0) { ?>
				<div class="ewPager ewRec">
					<span><?php echo $_SESSION['Language']->Phrase("Record") ?>&nbsp;<?php echo $a_customers_list->Pager->FromIndex ?>&nbsp;<?php echo $_SESSION['Language']->Phrase("To") ?>&nbsp;<?php echo $a_customers_list->Pager->ToIndex ?>&nbsp;<?php echo $_SESSION['Language']->Phrase("Of") ?>&nbsp;<?php echo $a_customers_list->Pager->RecordCount ?></span>
				</div>
		<?php } ?>

<?php if ($a_customers_list->TotalRecs > 0 && ((MS_SELECTABLE_PAGE_SIZES_POSITION=="Right" && $_SESSION['Language']->Phrase("dir")!="rtl") || (MS_SELECTABLE_PAGE_SIZES_POSITION=="Right" && $_SESSION['Language']->Phrase("dir")=="rtl"))) { ?>
<div class="ewPager"><span>&nbsp;<?php echo $_SESSION['Language']->Phrase("RecordsPerPage") ?>&nbsp;</span>
<input type="hidden" name="t" value="a_customers">
<select name="<?php echo EW_TABLE_REC_PER_PAGE ?>" class="form-control input-sm" onchange="this.form.submit();">
<option value="1"<?php if ($a_customers_list->DisplayRecs == 1) { ?> selected="selected"<?php } ?>>1</option>
<option value="3"<?php if ($a_customers_list->DisplayRecs == 3) { ?> selected="selected"<?php } ?>>3</option>
<option value="5"<?php if ($a_customers_list->DisplayRecs == 5) { ?> selected="selected"<?php } ?>>5</option>
<option value="10"<?php if ($a_customers_list->DisplayRecs == 10) { ?> selected="selected"<?php } ?>>10</option>
<option value="20"<?php if ($a_customers_list->DisplayRecs == 20) { ?> selected="selected"<?php } ?>>20</option>
<option value="50"<?php if ($a_customers_list->DisplayRecs == 50) { ?> selected="selected"<?php } ?>>50</option>
<option value="100"<?php if ($a_customers_list->DisplayRecs == 100) { ?> selected="selected"<?php } ?>>100</option>
</select>
</div>
<?php } // end if (MS_SELECTABLE_PAGE_SIZES_POSITION=="Right") ?>

</form>
<?php } ?>
<div class="ewListOtherOptions">
<?php
	foreach ($a_customers_list->OtherOptions as &$option)
		$option->RenderOptions("body", "bottom");
?>
</div>
<div class="clearfix"></div>
</div>
<?php } ?>
<?php } ?>
<?php // End of modification Customize Navigation/Pager Panel, by Masino Sinaga, May 2, 2012 ?>
</div>
<?php } ?>
<?php if (MS_SHOW_EMPTY_TABLE_ON_LIST_PAGE==FALSE) { ?>
<?php if ($a_customers_list->TotalRecs == 0 && $a_customers->CurrentAction == "") { // Show other options ?>
<div class="ewListOtherOptions">
<?php
	foreach ($a_customers_list->OtherOptions as &$option) {
		$option->ButtonClass = "";
		$option->RenderOptions("body", "");
	}
?>
</div>
<div class="clearfix"></div>
<?php } ?>
<?php } // MS_SHOW_EMPTY_TABLE_ON_LIST_PAGE is false ?>
<?php if ($a_customers->Export == "") { ?>
<script type="text/javascript">
fa_customerslistsrch.Init();
fa_customerslistsrch.FilterList = <?php echo $a_customers_list->GetFilterList() ?>;
fa_customerslist.Init();
</script>
<?php } ?>
<?php
$a_customers_list->ShowPageFooter();
if (EW_DEBUG_ENABLED)
	echo ew_DebugMsg();
?>
<?php if ($a_customers->Export == "") { ?>
<script type="text/javascript">
$(document).ready(function() {
	$("td:has(.ewAggregate)").css({"text-align": "right", "font-weight": "bold"}).find(".ewAggregate").hide();
});
$(document).on("preview", function(e, args) {
	var $tabpane = args.$tabpane;
	$tabpane.find("td:has(.ewAggregate)").css({"text-align": "right", "font-weight": "bold"}).find(".ewAggregate").hide();
});
</script>
<?php if (MS_USE_TABLE_SETTING_FOR_SEARCH_PANEL_COLLAPSED) { ?>
<?php if (isset($_SESSION['table_a_customers_views']) && $_SESSION['table_a_customers_views'] == 1) { ?>
	<?php if (CurrentPage()->SearchPanelCollapsed==FALSE) { ?>
<script type="text/javascript">
$(document).ready(function() {
	var SearchToggle = $('.ewSearchToggle'); var SearchPanel = $('.ewSearchPanel');
	SearchPanel.addClass('in'); SearchToggle.addClass('active');
});
</script>
	<?php } elseif (CurrentPage()->SearchPanelCollapsed==TRUE) { ?>
<script type="text/javascript">
$(document).ready(function() {
	var SearchToggle = $('.ewSearchToggle'); var SearchPanel = $('.ewSearchPanel');
	SearchPanel.removeClass('in'); SearchToggle.removeClass('active');
});
</script>

	<?php } ?>
<?php } else { ?>
<?php if (MS_USE_TABLE_SETTING_FOR_SEARCH_PANEL_STATUS==TRUE && MS_USE_PHPMAKER_SETTING_FOR_INITIATE_SEARCH_PANEL==TRUE) { ?>
<script type="text/javascript">
$(document).ready(function() { var SearchToggle = $('.ewSearchToggle'); var SearchPanel = $('.ewSearchPanel'); if(getCookie('a_customers_searchpanel')=="active"){ SearchToggle.addClass(getCookie('a_customers_searchpanel')); SearchPanel.addClass('in'); SearchToggle.addClass('active'); }else{ SearchPanel.removeClass('in'); SearchToggle.removeClass('active'); } SearchToggle.on('click',function(event) { event.preventDefault(); if (SearchToggle.hasClass('active')){ createCookie("a_customers_searchpanel", "notactive", 1); }else{ createCookie("a_customers_searchpanel", "active", 1); } }); });
</script>
<?php } elseif (MS_USE_TABLE_SETTING_FOR_SEARCH_PANEL_STATUS==TRUE && MS_USE_PHPMAKER_SETTING_FOR_INITIATE_SEARCH_PANEL==FALSE) { ?>
<script type="text/javascript">
$(document).ready(function() { var SearchToggle = $('.ewSearchToggle'); var SearchPanel = $('.ewSearchPanel'); if(getCookie('a_customers_searchpanel')=="active"){ SearchToggle.addClass(getCookie('a_customers_searchpanel')); SearchPanel.addClass('in'); SearchToggle.addClass('active'); }else{ SearchPanel.removeClass('in'); SearchToggle.removeClass('active'); } SearchToggle.on('click',function(event) { event.preventDefault(); if (SearchToggle.hasClass('active')){ createCookie("a_customers_searchpanel", "notactive", 1); }else{ createCookie("a_customers_searchpanel", "active", 1); } }); });
</script>
<?php } ?>
<?php } ?>
<?php } else { // end of MS_USE_TABLE_SETTING_FOR_SEARCH_PANEL_COLLAPSED ?>
<?php if (MS_USE_TABLE_SETTING_FOR_SEARCH_PANEL_STATUS==TRUE && MS_USE_PHPMAKER_SETTING_FOR_INITIATE_SEARCH_PANEL==TRUE) { ?>
	<?php if (isset($_SESSION['table_a_customers_views']) && $_SESSION['table_a_customers_views'] == 1) { ?>
<script type="text/javascript">
$(document).ready(function() { var SearchToggle = $('.ewSearchToggle'); var SearchPanel = $('.ewSearchPanel'); if(getCookie('a_customers_searchpanel')=="active"){ SearchToggle.addClass(getCookie('a_customers_searchpanel')); SearchPanel.addClass('in'); SearchToggle.addClass('active'); }else{ SearchPanel.removeClass('in'); SearchToggle.removeClass('active'); } SearchToggle.on('click',function(event) { event.preventDefault(); if (SearchToggle.hasClass('active')){ createCookie("a_customers_searchpanel", "notactive", 1); }else{ createCookie("a_customers_searchpanel", "active", 1); } }); });
</script>
	<?php } ?>
<?php } elseif (MS_USE_TABLE_SETTING_FOR_SEARCH_PANEL_STATUS==TRUE && MS_USE_PHPMAKER_SETTING_FOR_INITIATE_SEARCH_PANEL==FALSE) { ?>
<script type="text/javascript">
$(document).ready(function() { var SearchToggle = $('.ewSearchToggle'); var SearchPanel = $('.ewSearchPanel'); if(getCookie('a_customers_searchpanel')=="active"){ SearchToggle.addClass(getCookie('a_customers_searchpanel')); SearchPanel.addClass('in'); SearchToggle.addClass('active'); }else{ SearchPanel.removeClass('in'); SearchToggle.removeClass('active'); } SearchToggle.on('click',function(event) { event.preventDefault(); if (SearchToggle.hasClass('active')){ createCookie("a_customers_searchpanel", "notactive", 1); }else{ createCookie("a_customers_searchpanel", "active", 1); } }); });
</script>
<?php } ?>
<?php } ?>
<?php if (CurrentPage()->ListOptions->UseDropDownButton == TRUE) { ?>
<?php if (MS_USE_TABLE_SETTING_FOR_DROPUP_LISTOPTIONS == TRUE) { ?>
<script type="text/javascript">
$(document).ready(function() {
	var reccount = <?php echo CurrentPage()->RowCnt; ?>;
	var rowdropup = 4;
	if (reccount > 6) {
		for ( var i = 0; i <= (rowdropup - 1); i++ ) {
			$('#r' + (reccount - i) + '_<?php echo CurrentPage()->TableName; ?> .ewButtonDropdown').addClass('dropup');
		}
	}
});
</script>
<?php } ?>
<?php } ?>
<?php if ($a_customers->Export == "") { ?>
<script type="text/javascript">
$('.ewGridSave, .ewGridInsert').attr('onclick', 'return alertifySaveGrid(this)'); function alertifySaveGrid(obj) { <?php $_SESSION['Language']; ?> if (fa_customerslist.Validate() == true ) { alertify.confirm("<?php echo $_SESSION['Language']->Phrase('AlertifySaveGridConfirm'); ?>", function (e) { if (e) { $(window).unbind('beforeunload'); alertify.success("<?php echo $_SESSION['Language']->Phrase('AlertifySaveGrid'); ?>"); $("#fa_customerslist").submit(); } }).set("title", "<?php echo $_SESSION['Language']->Phrase('AlertifyConfirm'); ?>").set("defaultFocus", "cancel").set('oncancel', function(closeEvent){ alertify.error('<?php echo $_SESSION['Language']->Phrase('AlertifyCancel'); ?>');}).set('labels', {ok:'<?php echo $_SESSION['Language']->Phrase("MyOKMessage"); ?>!', cancel:'<?php echo $_SESSION['Language']->Phrase("MyCancelMessage"); ?>'}); } return false; }
</script>
<script type="text/javascript">
$('.ewInlineUpdate').attr('onclick', 'return alertifySaveInlineEdit(this)'); function alertifySaveInlineEdit(obj) { <?php $_SESSION['Language']; ?> if (fa_customerslist.Validate() == true ) { alertify.confirm("<?php echo $_SESSION['Language']->Phrase('AlertifySaveGridConfirm'); ?>", function (e) { if (e) { $(window).unbind('beforeunload'); alertify.success("<?php echo $_SESSION['Language']->Phrase('AlertifySaveGrid'); ?>"); $("#fa_customerslist").submit(); } }).set("title", "<?php echo $_SESSION['Language']->Phrase('AlertifyConfirm'); ?>").set("defaultFocus", "cancel").set('oncancel', function(closeEvent){ alertify.error('<?php echo $_SESSION['Language']->Phrase('AlertifyCancel'); ?>');}).set('labels', {ok:'<?php echo $_SESSION['Language']->Phrase("MyOKMessage"); ?>!', cancel:'<?php echo $_SESSION['Language']->Phrase("MyCancelMessage"); ?>'}); } return false; }
</script>
<script type="text/javascript">
$('.ewInlineInsert').attr('onclick', 'return alertifySaveInlineInsert(this)'); function alertifySaveInlineInsert(obj) { <?php  $_SESSION['Language']; ?> if (fa_customerslist.Validate() == true ) { alertify.confirm("<?php echo $_SESSION['Language']->Phrase('AlertifySaveGridConfirm'); ?>", function (e) { if (e) { $(window).unbind('beforeunload'); alertify.success("<?php echo $_SESSION['Language']->Phrase('AlertifySaveGrid'); ?>"); $("#fa_customerslist").submit(); } }).set("title", "<?php echo $_SESSION['Language']->Phrase('AlertifyConfirm'); ?>").set("defaultFocus", "cancel").set('oncancel', function(closeEvent){ alertify.error('<?php echo $_SESSION['Language']->Phrase('AlertifyCancel'); ?>');}).set('labels', {ok:'<?php echo $_SESSION['Language']->Phrase("MyOKMessage"); ?>!', cancel:'<?php echo $_SESSION['Language']->Phrase("MyCancelMessage"); ?>'}); } return false; }
</script>
<?php } ?>
<?php if ($a_customers->CurrentAction == "" || $a_customers->Export == "") { // Change && become || in order to add scroll table in Grid, by Masino Sinaga, August 3, 2014 ?>
<script type="text/javascript">
<?php if (MS_TABLE_WIDTH_STYLE==1) { // Begin of modification Optimizing Main Table Width to Maximum Width of Site, by Masino Sinaga, April 30, 2012 ?>
<?php $iWidthAdjustment = (MS_MENU_HORIZONTAL) ? 0 : 100; ?>
ew_ScrollableTable("gmp_a_customers", "<?php echo (MS_SCROLL_TABLE_WIDTH - $iWidthAdjustment); ?>px", "<?php echo MS_SCROLL_TABLE_HEIGHT; ?>px");
ew_ScrollableTable("gmp_a_customers_empty_table", "<?php echo (MS_SCROLL_TABLE_WIDTH - $iWidthAdjustment); ?>px", "<?php echo MS_SCROLL_TABLE_HEIGHT; ?>px");
<?php } elseif (MS_TABLE_WIDTH_STYLE==2) { ?>
ew_ScrollableTable("gmp_a_customers", "<?php echo MS_SCROLL_TABLE_WIDTH; ?>px", "<?php echo MS_SCROLL_TABLE_HEIGHT; ?>px");
ew_ScrollableTable("gmp_a_customers_empty_table", "<?php echo MS_SCROLL_TABLE_WIDTH; ?>px", "<?php echo MS_SCROLL_TABLE_HEIGHT; ?>px");
<?php } elseif (MS_TABLE_WIDTH_STYLE==3) { ?>
ew_ScrollableTable("gmp_a_customers", "100%", "<?php echo MS_SCROLL_TABLE_HEIGHT; ?>px");
ew_ScrollableTable("gmp_a_customers_empty_table", "100%", "<?php echo MS_SCROLL_TABLE_HEIGHT; ?>px");
<?php } // End of modification Optimizing Main Table Width to Maximum Width of Site, by Masino Sinaga, April 30, 2012 ?>
<?php } ?>
</script>
<?php } ?>
<?php include_once "footer.php" ?>
<?php
$a_customers_list->Page_Terminate();
?>
