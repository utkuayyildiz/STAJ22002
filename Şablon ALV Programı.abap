REPORT Z_DENEME_ALV.

*&---------------------------------------------------------------------*
*& CLASS lcl_alv DEFINITION
*&---------------------------------------------------------------------*
CLASS lcl_alv DEFINITION DEFERRED.

DATA: go_alv TYPE REF TO lcl_alv.
PARAMETERS: Tablo TYPE dd02l-tabname.
DATA: lv_size TYPE so_obj_len.

*&---------------------------------------------------------------------*
*& CLASS lcl_eh_head DEFINITION
*&---------------------------------------------------------------------*
CLASS lcl_eh_head DEFINITION.
  PUBLIC SECTION.
    DATA rpt_ref TYPE REF TO lcl_alv.
    METHODS constructor IMPORTING i_grid TYPE REF TO cl_gui_alv_grid i_rpt TYPE REF TO lcl_alv.
  PRIVATE SECTION.
    DATA grid TYPE REF TO cl_gui_alv_grid.
    METHODS on_toolbar       FOR EVENT toolbar       OF cl_gui_alv_grid IMPORTING e_object e_interactive.
    METHODS on_top_of_page   FOR EVENT top_of_page   OF cl_gui_alv_grid IMPORTING e_dyndoc_id.
ENDCLASS.

*&---------------------------------------------------------------------*
*& CLASS lcl_alv DEFINITION
*&---------------------------------------------------------------------*
CLASS lcl_alv DEFINITION.
  PUBLIC SECTION.
    DATA: mt_outtab   TYPE REF TO data,
          mv_tabname  TYPE dd02l-tabname,
          go_grd      TYPE REF TO cl_gui_alv_grid,
          gt_fct      TYPE lvc_t_fcat,
          gs_lay      TYPE lvc_s_layo.

    METHODS constructor  IMPORTING i_str TYPE dd02l-tabname.
    METHODS get_data.
    METHODS send_report_via_email.
    METHODS cre_fct.
    METHODS cre_lay.
    METHODS dis_alv.
ENDCLASS.

*&---------------------------------------------------------------------*
*& CLASS lcl_eh_head IMPLEMENTATION
*&---------------------------------------------------------------------*
CLASS lcl_eh_head IMPLEMENTATION.
  METHOD constructor.
    grid = i_grid. rpt_ref = i_rpt.
    SET HANDLER on_toolbar FOR grid.
    SET HANDLER on_top_of_page FOR grid.
    grid->register_edit_event( i_event_id = cl_gui_alv_grid=>mc_evt_modified ).
  ENDMETHOD.
  METHOD on_toolbar. APPEND VALUE #( butn_type = 3 ) TO e_object->mt_toolbar. ENDMETHOD.
  METHOD on_top_of_page. CALL METHOD e_dyndoc_id->add_picture EXPORTING picture_id = 'FATURALAB'. ENDMETHOD.
ENDCLASS.

*&---------------------------------------------------------------------*
*& CLASS lcl_alv IMPLEMENTATION
*&---------------------------------------------------------------------*
CLASS lcl_alv IMPLEMENTATION.
  METHOD constructor.
    mv_tabname = i_str.
    CREATE DATA mt_outtab TYPE STANDARD TABLE OF (mv_tabname) WITH EMPTY KEY.
  ENDMETHOD.

  METHOD get_data.
    FIELD-SYMBOLS <lt_data> TYPE INDEX TABLE.
    ASSIGN mt_outtab->* TO <lt_data>.
    SELECT * FROM (mv_tabname)
 INTO TABLE @<lt_data> UP TO 250 ROWS.
  ENDMETHOD.

  METHOD cre_fct.
    CALL FUNCTION 'LVC_FIELDCATALOG_MERGE'
      EXPORTING i_structure_name = mv_tabname
      CHANGING  ct_fieldcat      = gt_fct.
  ENDMETHOD.

  METHOD cre_lay.
    gs_lay = VALUE #( zebra = 'X' sel_mode = 'A' cwidth_opt = 'X' ).
  ENDMETHOD.

  METHOD send_report_via_email.
    DATA: lv_content       TYPE string,
          lv_binary        TYPE solix_tab,
          lv_size          TYPE so_obj_len,
          lo_send_request  TYPE REF TO cl_bcs,
          lo_document      TYPE REF TO cl_document_bcs,
          lo_recipient     TYPE REF TO if_recipient_bcs,
          lx_bcs_exception TYPE REF TO cx_bcs,
          lt_fcat          TYPE lvc_t_fcat.

    FIELD-SYMBOLS <lt_data> TYPE INDEX TABLE.
    ASSIGN mt_outtab->* TO <lt_data>.

    CALL FUNCTION 'LVC_FIELDCATALOG_MERGE'
      EXPORTING i_structure_name = mv_tabname
      CHANGING  ct_fieldcat      = lt_fcat.

    LOOP AT lt_fcat INTO DATA(ls_fcat).
      lv_content = COND #( WHEN lv_content IS INITIAL THEN ls_fcat-fieldname ELSE |{ lv_content };{ ls_fcat-fieldname }| ).
    ENDLOOP.
    lv_content = lv_content && cl_abap_char_utilities=>newline.

    LOOP AT <lt_data> ASSIGNING FIELD-SYMBOL(<fs_line>).
      DATA(lv_row) = VALUE string( ).
      DO.
        ASSIGN COMPONENT sy-index OF STRUCTURE <fs_line> TO FIELD-SYMBOL(<fs_val>).
        IF sy-subrc <> 0. EXIT. ENDIF.
        DATA(lv_val_str) = |{ <fs_val> }|. CONDENSE lv_val_str.
        lv_row = COND #( WHEN lv_row IS INITIAL THEN lv_val_str ELSE |{ lv_row };{ lv_val_str }| ).
      ENDDO.
      lv_content = lv_content && lv_row && cl_abap_char_utilities=>newline.
    ENDLOOP.

    TRY.
        cl_bcs_convert=>string_to_solix(
          EXPORTING iv_string = lv_content iv_codepage = '4110' iv_add_bom = abap_true
          IMPORTING et_solix = lv_binary ev_size = lv_size ).

        lo_send_request = cl_bcs=>create_persistent( ).
        lo_document = cl_document_bcs=>create_document(
                        i_type = 'RAW'
                        i_text = VALUE #( ( line = 'Günlük Rapor Ektedir.' ) )
                        i_subject = |Otomatik fatura tablosu | ).

        lo_document->add_attachment( i_attachment_type = 'XLS' i_attachment_subject = 'Rapor'
                                     i_attachment_size = lv_size i_att_content_hex = lv_binary ).

        lo_send_request->set_document( lo_document ).
        lo_recipient = cl_cam_address_bcs=>create_internet_address( 'utku.ayyildiz@faturalab.com' ).
        lo_send_request->add_recipient( lo_recipient ).
        lo_send_request->send( ).
        COMMIT WORK.
      CATCH cx_bcs INTO lx_bcs_exception.
        ROLLBACK WORK.
    ENDTRY.
  ENDMETHOD.

  METHOD dis_alv.
    FIELD-SYMBOLS <lt_data> TYPE INDEX TABLE.
    ASSIGN mt_outtab->* TO <lt_data>.
    go_grd->set_table_for_first_display( EXPORTING is_layout = gs_lay CHANGING it_outtab = <lt_data> it_fieldcatalog = gt_fct ).
  ENDMETHOD.
ENDCLASS.

*&---------------------------------------------------------------------*
*& START-OF-SELECTION
*&---------------------------------------------------------------------*
START-OF-SELECTION.
  go_alv = NEW lcl_alv( i_str = Tablo ).
  go_alv->get_data( ).

  IF sy-batch = 'X'.
    go_alv->send_report_via_email( ).
  ELSE.
    go_alv->cre_fct( ).
    go_alv->cre_lay( ).
    CALL SCREEN 100.
  ENDIF.

*&---------------------------------------------------------------------*
*& SCREEN 100 MODULES
*&---------------------------------------------------------------------*
MODULE status_0100 OUTPUT.
  SET PF-STATUS '100'.
  IF go_alv IS BOUND AND go_alv->go_grd IS NOT BOUND.
    DATA(lo_cont) = NEW cl_gui_custom_container( container_name = 'CUSTOM' ).
    CREATE OBJECT go_alv->go_grd EXPORTING i_parent = lo_cont.
    DATA(lo_eh) = NEW lcl_eh_head( i_grid = go_alv->go_grd i_rpt = go_alv ).
    go_alv->dis_alv( ).
  ENDIF.
ENDMODULE.

MODULE user_command_0100 INPUT.
  IF sy-ucomm = '&BACK' OR sy-ucomm = '&EXIT' OR sy-ucomm = '&CANCEL'. LEAVE TO SCREEN 0. ENDIF.
ENDMODULE.