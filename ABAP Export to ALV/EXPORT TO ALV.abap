*&---------------------------------------------------------------------*
*& Report ZEXPORT_TO_EXCEL
*&---------------------------------------------------------------------*
*& Export internal table data to XLSX Excel format with proper style
*&---------------------------------------------------------------------*
*& Author: Nikhil Mihra
*& Date  : 12.09.2020
*&---------------------------------------------------------------------*
*& 1. Fetch data into an internal table
*& 2. Prepare fieldcatalog if not already
*& 3. Merge data table and fieldcatalog into one
*& 4. Convert the internal tabe data to XSTRING format.
*& 5. Convert XSTRING to BINARY format
*& 6. Just pass it to GUI_DOWNLOAD
*&  OR
*& Just copy this code and modify
*&---------------------------------------------------------------------*
REPORT ZEXPORT_TO_EXCEL.

" 1. First capture the data into an internal table

DATA(DATUM) = SY-DATUM - 180.

SELECT VBELN, ERDAT, VBTYP, AUART, VKORG, VTWEG, SPART
         FROM VBAK
          INTO TABLE @DATA(LT_VBAK)
            WHERE ERDAT BETWEEN @DATUM AND @SY-DATUM.

" 2. We need fieldcatalog and we are gonna do that OO ABAP way.

TRY.
CL_SALV_TABLE=>FACTORY(
  EXPORTING
    LIST_DISPLAY   = IF_SALV_C_BOOL_SAP=>FALSE    " ALV DISPLAYED IN LIST MODE
*    R_CONTAINER    =     " PASS CONTAINER OBJECTS ONLY IF YOU HAVE USED CONTAINER IN YOUR PROGRAM
*    CONTAINER_NAME =
  IMPORTING
    R_SALV_TABLE   = DATA(LO_SALV_TABLE)    " YOU CAN DO THIS ONLY IN CLASS METHODS, YUP THAT'S SAD
  CHANGING
    T_TABLE        = LT_VBAK   " YUP, THAT'S THE ONLY THING IT NEEDS
).
  CATCH CX_SALV_MSG.
ENDTRY.

DATA(LT_FIELDCAT) = CL_SALV_CONTROLLER_METADATA=>GET_LVC_FIELDCATALOG(
                    EXPORTING
                      R_COLUMNS      = LO_SALV_TABLE->GET_COLUMNS( )    " ALV Filter
                      R_AGGREGATIONS = LO_SALV_TABLE->GET_AGGREGATIONS( )    " ALV Aggregations

).

" 3. Now, its time to merge fieldcatalog and data table into one

" Don't worry, we method for that too
DATA LT_DATA TYPE REF TO DATA.

GET REFERENCE OF LT_VBAK INTO LT_DATA.

CL_SALV_EX_UTIL=>FACTORY_RESULT_DATA_TABLE(
  EXPORTING
    R_DATA                 = LT_DATA   " Data table
*    S_LAYOUT               =     " ALV Control: Layout Structure
    T_FIELDCATALOG         = LT_FIELDCAT    " Field Catalog for List Viewer Control
*    T_SORT                 =     " ALV Control: Table of Sort Criteria
  RECEIVING
    R_RESULT_DATA_TABLE    = DATA(LO_DATA)
).

" 4. Lets convert this into XSTRING format.

CL_SALV_BS_TT_UTIL=>IF_SALV_BS_TT_UTIL~TRANSFORM(
  EXPORTING
    XML_VERSION   = IF_SALV_BS_XML=>VERSION    " XML Version to be Selected
    GUI_TYPE      = IF_SALV_BS_XML=>c_gui_type_gui    " Constant
    XML_TYPE      = IF_SALV_BS_XML=>C_TYPE_XLSX    " !! Important
    XML_FLAVOUR   = IF_SALV_BS_C_TT=>C_TT_XML_FLAVOUR_FULL
    R_RESULT_DATA = LO_DATA
  IMPORTING
    XML           = DATA(XSTRING)
).

" 5. Convert XSTRING to BINARY format
DATA: LENGTH TYPE I,
      BINARY TYPE SOLIX_TAB.

CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
  EXPORTING
    BUFFER          = XSTRING
*    APPEND_TO_TABLE = SPACE
  IMPORTING
    OUTPUT_LENGTH   = LENGTH
  TABLES
    BINARY_TAB      = BINARY
  .

" 6. Now pass length and binary into GUI_DOWNLOAD
CL_GUI_FRONTEND_SERVICES=>GUI_DOWNLOAD(
  EXPORTING
    bin_filesize              = LENGTH    " File length for binary files
    filename                  = 'D://TEMP/EXPORT.XLSX'    " Name of file
    filetype                  = 'BIN'    " File type (ASCII, binary ...)
  CHANGING
    data_tab                  = BINARY    " Transfer table
  EXCEPTIONS
    file_write_error          = 1
    no_batch                  = 2
    gui_refuse_filetransfer   = 3
    invalid_type              = 4
    no_authority              = 5
    unknown_error             = 6
    header_not_allowed        = 7
    separator_not_allowed     = 8
    filesize_not_allowed      = 9
    header_too_long           = 10
    dp_error_create           = 11
    dp_error_send             = 12
    dp_error_write            = 13
    unknown_dp_error          = 14
    access_denied             = 15
    dp_out_of_memory          = 16
    disk_full                 = 17
    dp_timeout                = 18
    file_not_found            = 19
    dataprovider_exception    = 20
    control_flush_error       = 21
    not_supported_by_gui      = 22
    error_no_gui              = 23
    others                    = 24
).
IF sy-subrc <> 0.
 MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
            WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
ENDIF.
