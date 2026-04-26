REPORT Z_BATCH.

CLASS lcl_alv DEFINITION DEFERRED.

DATA: lv_subrc    TYPE sy-subrc,
      lt_messtab  TYPE TABLE OF bdcmsgcoll.
data: gv_datum type datum.

TABLES: lfa1.
SELECT-OPTIONS: s_lifnr FOR lfa1-lifnr NO-DISPLAY.
TYPES: BEGIN OF ty_balance,
         lifnr TYPE lifnr,
         name1 TYPE name1,
         t_girdi TYPE bseg-wrbtr,
         t_cikti TYPE bseg-wrbtr,
         bakiye  TYPE bseg-wrbtr,
         waers   TYPE waers,
         t_color TYPE lvc_t_scol,
       END OF ty_balance.

DATA: lt_balance TYPE TABLE OF ty_balance.
data: ls_balance type ty_balance.

DATA: gv_subscreen TYPE sy-dynnr.
DATA: bldat       TYPE bldat.
DATA: go_alv      TYPE REF TO lcl_alv.
DATA: gv_file2 type localfile.

DATA: bukrs TYPE bukrs,
      blart TYPE blart,
      dat   TYPE datum,
      monat TYPE monat,
      waers TYPE waers,
      amo   TYPE bseg-wrbtr,
      newko TYPE bseg-hkont,
      newk2 TYPE bseg-hkont,
      sgtxt TYPE sgtxt.

TYPES: BEGIN OF ty_excel,
         bukrs TYPE bukrs,
         blart TYPE blart,
         dat   TYPE datum,
         monat TYPE monat,
         waers TYPE waers,
         amo   TYPE char15,
         newko TYPE bseg-hkont,
         newk2 TYPE bseg-hkont,
       END OF ty_excel.

TYPES: BEGIN OF ty_outtab.
        INCLUDE TYPE zfat.
TYPES:  style TYPE lvc_t_styl,
       END OF ty_outtab.

DATA: rb1 TYPE char1,
      rb2 TYPE char1.

DATA: lt_excel TYPE TABLE OF ty_excel,
      ls_excel TYPE ty_excel.
DATA: lt_raw   TYPE truxs_t_text_data.
DATA: gv_file  TYPE localfile.

DATA: lv_bukrs_bdc TYPE bdc_fval,
      lv_blart_bdc TYPE bdc_fval,
      lv_dat_bdc   TYPE bdc_fval,
      lv_monat_bdc TYPE bdc_fval,
      lv_waers_bdc TYPE bdc_fval,
      lv_amo_bdc   TYPE bdc_fval,
      lv_newko_bdc TYPE bdc_fval,
      lv_newk2_bdc TYPE bdc_fval,
      lv_sgtxt_bdc TYPE bdc_fval.

PARAMETERS: p_tablo TYPE dd02l-tabname NO-DISPLAY.


CLASS lcl_eh_head DEFINITION.
  PUBLIC SECTION.
    DATA rpt_ref TYPE REF TO lcl_alv.
    METHODS constructor IMPORTING i_grid TYPE REF TO cl_gui_alv_grid i_rpt TYPE REF TO lcl_alv.
  PRIVATE SECTION.
    DATA grid TYPE REF TO cl_gui_alv_grid.
    METHODS on_toolbar       FOR EVENT toolbar       OF cl_gui_alv_grid IMPORTING e_object e_interactive.
    METHODS on_user_command   FOR EVENT user_command   OF cl_gui_alv_grid IMPORTING e_ucomm.
    METHODS on_hotspot_click FOR EVENT hotspot_click OF cl_gui_alv_grid IMPORTING e_row_id e_column_id es_row_no.
    METHODS on_button_click  FOR EVENT button_click  OF cl_gui_alv_grid IMPORTING es_col_id es_row_no.
    METHODS on_top_of_page   FOR EVENT top_of_page   OF cl_gui_alv_grid IMPORTING e_dyndoc_id.
ENDCLASS.

CLASS lcl_alv DEFINITION.
  PUBLIC SECTION.
    DATA: mt_outtab   TYPE REF TO data,
          mv_tabname  TYPE dd02l-tabname,
          go_grd      TYPE REF TO cl_gui_alv_grid,
          gt_fct      TYPE lvc_t_fcat,
          gs_lay      TYPE lvc_s_layo,
        go_dyndoc    TYPE REF TO cl_dd_document, " Hata: LO_DYNDOC unknown çözümü
          go_cont_logo TYPE REF TO cl_gui_container.
    TYPES: BEGIN OF ty_mess_pop,
             statu TYPE icon_d, hata TYPE bapi_msg,
           END OF ty_mess_pop,
           ty_mess_pop_t TYPE STANDARD TABLE OF ty_mess_pop WITH EMPTY KEY.

    CLASS-METHODS create IMPORTING i_str TYPE dd02l-tabname RETURNING VALUE(e_alv) TYPE REF TO lcl_alv.
    METHODS constructor  IMPORTING i_str TYPE dd02l-tabname.
    METHODS get_data.
    METHODS cre_fct.
    METHODS cre_lay.
    METHODS dis_alv.
    METHODS set_alv_hot   IMPORTING i_fnam TYPE string.
    METHODS disp_mess_popup IMPORTING it_mess TYPE ty_mess_pop_t.
    METHODS get_logo_url  RETURNING VALUE(r_url) TYPE char255.
ENDCLASS.

INITIALIZATION.
perform set_ini.



START-OF-SELECTION.
  gv_subscreen = '0101'.
  CALL SCREEN 100.

FORM z_excel.
  DATA: lv_belnr TYPE belnr_d,
        ls_zfat  TYPE zfat.

  CALL FUNCTION 'TEXT_CONVERT_XLS_TO_SAP'
    EXPORTING
      i_line_header        = 'X'
      i_filename           = gv_file
      i_tab_raw_data       = lt_raw
    TABLES
      i_tab_converted_data = lt_excel
    EXCEPTIONS
      conversion_failed    = 1
      OTHERS               = 2.

  IF sy-subrc <> 0.
    MESSAGE 'Excel okunamadı.' TYPE 'E'.
    RETURN.
  ENDIF.

  IF go_alv IS NOT BOUND.
    go_alv = lcl_alv=>create( i_str = 'ZFAT' ).
  ENDIF.
  " ---------------------------

  LOOP AT lt_excel INTO ls_excel.
    AUTHORITY-CHECK OBJECT 'F_BKPF_BUK'
      ID 'BUKRS' FIELD ls_excel-bukrs
      ID 'ACTVT' FIELD '01'.

    IF sy-subrc <> 0.
      MESSAGE 'Şirket kodu yetkiniz bulunmamaktadır'
              TYPE 'E'.
      RETURN.
    ENDIF.
    CLEAR: lv_dat_bdc, lv_belnr, lt_messtab.

    WRITE ls_excel-dat TO lv_dat_bdc.
    CONDENSE lv_dat_bdc NO-GAPS.

    WRITE ls_excel-blart TO lv_blart_bdc.
    lv_bukrs_bdc = ls_excel-bukrs.
    WRITE ls_excel-dat   TO lv_dat_bdc.
    WRITE ls_excel-monat TO lv_monat_bdc.
    lv_monat_bdc = ls_excel-monat.
    WRITE ls_excel-waers TO lv_waers_bdc.
    lv_waers_bdc = ls_excel-waers.
      WRITE ls_excel-amo TO lv_amo_bdc.
    CONDENSE lv_amo_bdc NO-GAPS.
    WRITE ls_excel-newko TO lv_newko_bdc.
    lv_newko_bdc = ls_excel-newko.
    WRITE ls_excel-newk2 TO lv_newk2_bdc.
    lv_newk2_bdc = ls_excel-newk2.


       CALL FUNCTION 'ZF_02'
      EXPORTING
        ctu       = 'X'
        mode      = 'N'
        update    = 'S'
        nodata    = '/'
        bldat_001 = lv_dat_bdc
        blart_002 = lv_blart_bdc
        bukrs_003 = lv_bukrs_bdc
        budat_004 = lv_dat_bdc
        monat_005 = lv_monat_bdc
        waers_006 = lv_waers_bdc
        wwert_007 = lv_dat_bdc
        newko_010 = lv_newko_bdc
        wrbtr_011 = lv_amo_bdc
        wrbtr_018 = lv_amo_bdc
        wrbtr_015 = lv_amo_bdc
        newko_013 = lv_newk2_bdc

      IMPORTING

        subrc     = lv_subrc

      TABLES

        messtab   = lt_messtab.

            IF lv_subrc EQ 0.
              READ TABLE lt_messtab INTO DATA(ls_msg) WITH KEY msgid = 'F5' msgnr = '312'.
              IF sy-subrc = 0.
                lv_belnr = ls_msg-msgv1.

                IF lv_belnr IS NOT INITIAL.
                  CLEAR ls_zfat.
                  ls_zfat-mandt = sy-mandt.
                  ls_zfat-bukrs = ls_excel-bukrs.
                  ls_zfat-belnr = lv_belnr.
                  ls_zfat-gjahr = ls_excel-dat(4).
                  ls_zfat-blart = ls_excel-blart.
                  ls_zfat-monat = ls_excel-monat.
                  ls_zfat-waers = ls_excel-waers.
                  ls_zfat-amo   = ls_excel-amo.
                  ls_zfat-uname = sy-uname.
                  ls_zfat-cpudt = sy-datum.
                  ls_zfat-cputm = sy-uzeit.
                  IF ls_zfat-gjahr ne '2026'.
                      MESSAGE 'Mali yıl 2026dan önce olamaz' type 'I'.
                      CONTINUE.
                  ENDIF.

                  MODIFY zfat FROM ls_zfat.
                  COMMIT WORK AND WAIT.
                ENDIF.
              ENDIF.
            ENDIF.

  ENDLOOP.

DATA: lt_display_mess TYPE lcl_alv=>ty_mess_pop_t,
      ls_display_mess TYPE lcl_alv=>ty_mess_pop.
DATA: lv_msg_text(200) TYPE c.

IF lt_messtab IS NOT INITIAL.
  LOOP AT lt_messtab INTO DATA(ls_bdc_msg).
    CHECK ls_bdc_msg-msgtyp = 'E' OR ls_bdc_msg-msgtyp = 'A'.

    CLEAR lv_msg_text.
    MESSAGE ID ls_bdc_msg-msgid
            TYPE ls_bdc_msg-msgtyp
            NUMBER ls_bdc_msg-msgnr
            WITH ls_bdc_msg-msgv1 ls_bdc_msg-msgv2 ls_bdc_msg-msgv3 ls_bdc_msg-msgv4
            INTO lv_msg_text.

    ls_display_mess-statu = icon_led_red. " Hata ikonu
    ls_display_mess-hata  = lv_msg_text.
    APPEND ls_display_mess TO lt_display_mess.
  ENDLOOP.

  IF lt_display_mess IS NOT INITIAL.
    go_alv->disp_mess_popup( it_mess = lt_display_mess ).
  ENDIF.
ENDIF.

IF go_alv IS NOT BOUND.
    go_alv = lcl_alv=>create( i_str = 'ZFAT' ).
  ENDIF.

  go_alv->get_data( ).
  CALL SCREEN 103.

ENDFORM.

CLASS lcl_eh_head IMPLEMENTATION.
  METHOD constructor.
    grid = i_grid. rpt_ref = i_rpt.
    SET HANDLER on_toolbar FOR grid.
    SET HANDLER on_user_command FOR grid.
    SET HANDLER on_hotspot_click FOR grid.
    SET HANDLER on_button_click FOR grid.
    SET HANDLER on_top_of_page FOR grid.
    grid->register_edit_event( i_event_id = cl_gui_alv_grid=>mc_evt_modified ).
  ENDMETHOD.

  METHOD on_toolbar.
    APPEND VALUE #( butn_type = 3 ) TO e_object->mt_toolbar.
    APPEND VALUE #( function = 'DB'
                    quickinfo = 'Değişiklikleri Kaydet'
                    text = 'Verileri Kaydet' ) to e_object->mt_toolbar.
    APPEND VALUE #( function = 'DELETE'
                    quickinfo = 'Seçili satırları veritabanından sil'
                    text = 'Satır sil'
                    icon = icon_delete_row ) to e_object->mt_toolbar.
    append value #( function = 'ADD'
                    quickinfo = 'Yeni veri ekle'
                    text = 'Satır ekle'
                    icon = icon_insert_row ) to e_object->mt_toolbar.

  ENDMETHOD.

METHOD on_user_command.
  FIELD-SYMBOLS <lt_outtab> TYPE INDEX TABLE.
  DATA: lt_messtab_alv TYPE TABLE OF bdcmsgcoll,
        lv_subrc_alv   TYPE sy-subrc,
        ls_zfat_old    TYPE zfat,
        ls_zfat        TYPE zfat,
        lv_dat_bdc     TYPE bdc_fval,
        lv_amo_bdc     TYPE bdc_fval.

data: lt_rows type lvc_t_row,
      ls_row type lvc_s_row.

  CASE e_ucomm.
   WHEN 'DB'.
      go_alv->go_grd->check_changed_data( ).
      ASSIGN go_alv->mt_outtab->* TO <lt_outtab>.

      go_alv->go_grd->get_selected_rows( IMPORTING et_index_rows = lt_rows ).

      IF lt_rows IS INITIAL.
        MESSAGE 'Lütfen en az bir satır seçiniz.' TYPE 'I'.
        RETURN.
      ENDIF.

      LOOP AT lt_rows INTO ls_row.
        READ TABLE <lt_outtab> ASSIGNING FIELD-SYMBOL(<ls_data>) INDEX ls_row-index.
        CHECK sy-subrc = 0.

        MOVE-CORRESPONDING <ls_data> TO ls_zfat.

        SELECT SINGLE * FROM zfat INTO ls_zfat_old
          WHERE bukrs = ls_zfat-bukrs
            AND belnr = ls_zfat-belnr
            AND gjahr = ls_zfat-gjahr.

        IF ls_zfat <> ls_zfat_old OR ls_zfat-belnr IS INITIAL.

          CLEAR: lv_dat_bdc, lv_amo_bdc.

          WRITE ls_zfat-cpudt TO lv_dat_bdc.
          CONDENSE lv_dat_bdc NO-GAPS.
          WRITE ls_zfat-amo TO lv_amo_bdc.
          CONDENSE lv_amo_bdc NO-GAPS.

          CALL FUNCTION 'ZF_02'
            EXPORTING
              ctu       = 'X'
              mode      = 'N'
              update    = 'S'
              bldat_001 = lv_dat_bdc
              blart_002 = CONV bdc_fval( ls_zfat-blart )
              bukrs_003 = CONV bdc_fval( ls_zfat-bukrs )
              budat_004 = lv_dat_bdc
              monat_005 = CONV bdc_fval( ls_zfat-monat )
              waers_006 = CONV bdc_fval( ls_zfat-waers )
              newko_010 = '100000000'
              wrbtr_011 = lv_amo_bdc
              newko_013 = '100000000'
              wrbtr_015 = lv_amo_bdc
              wrbtr_018 = lv_amo_bdc
            IMPORTING
              subrc     = lv_subrc_alv
            TABLES
              messtab   = lt_messtab_alv.

          IF lv_subrc_alv = 0.
            READ TABLE lt_messtab_alv INTO DATA(ls_msg)
                 WITH KEY msgid = 'F5' msgnr = '312'.
            IF sy-subrc = 0.
              ls_zfat-belnr = ls_msg-msgv1.
              ls_zfat-uname = sy-uname.
              ls_zfat-cpudt = sy-datum.
              ls_zfat-cputm = sy-uzeit.
              MODIFY zfat FROM ls_zfat.

              MOVE-CORRESPONDING ls_zfat TO <ls_data>.
            ENDIF.
          ENDIF.
        ENDIF.

        CLEAR: lt_messtab_alv, lv_subrc_alv, ls_zfat, ls_zfat_old.
      ENDLOOP.

      COMMIT WORK AND WAIT.

      go_alv->go_grd->refresh_table_display( ).

      MESSAGE 'İşlem başarıyla tamamlandı ve ZFAT tablosu güncellendi.' TYPE 'S'.

      when 'ADD'.
        assign go_alv->mt_outtab->* to <lt_outtab>.
        append INITIAL LINE TO <lt_outtab> ASSIGNING FIELD-SYMBOL(<ls_new>).
        assign COMPONENT 'NEWKO' of STRUCTURE <ls_new> to FIELD-SYMBOL(<lv_newko>).
        IF <lv_newko> is ASSIGNED.
          <lv_newko> = '100000000'.
        ENDIF.
        assign COMPONENT 'NEWK2' OF STRUCTURE <ls_new> to FIELD-SYMBOL(<lv_newk2>).
        IF <lv_newk2> is ASSIGNED.
          <lv_newk2> = '100000000'.
        ENDIF.
        go_alv->go_grd->refresh_table_display( ).
      when 'DELETE'.
        go_alv->go_grd->get_selected_rows(
          IMPORTING
            et_index_rows =      LT_ROWS
        ).

  IF lt_rows is initial.
    message 'Lütfen satır seçiniz' type 'E'.
    return.
  ENDIF.
assign go_alv->mt_outtab->* to <lt_outtab>.
sort lt_rows by index DESCENDING.

  LOOP AT lt_rows into ls_row.
    READ TABLE <lt_outtab> ASSIGNING FIELD-SYMBOL(<ls_del_row>) index ls_row-index.
    IF sy-subrc eq 0.
      MOVE-CORRESPONDING <ls_del_row> to ls_zfat.
    ENDIF.

    delete from zfat where bukrs eq ls_zfat-bukrs
                     and belnr eq ls_zfat-belnr.

  IF sy-subrc eq 0.
    delete <lt_outtab> index ls_row-index.
  ENDIF.
  ENDLOOP.
  COMMIT WORK and wait.
  message 'Kayıtlar silindi' type 'S'.

      go_alv->go_grd->refresh_table_display( ).
  ENDCASE.
ENDMETHOD.
METHOD on_hotspot_click.
    DATA: ls_row_data TYPE zfat.
    DATA: lr_data     TYPE REF TO data.

    lr_data = rpt_ref->mt_outtab.
    CASE e_column_id-fieldname.
      WHEN 'BELNR'.
        IF es_row_no-row_id IS NOT INITIAL.

          FIELD-SYMBOLS <lt_table> TYPE INDEX TABLE.
          ASSIGN lr_data->* TO <lt_table>.

          READ TABLE <lt_table> INDEX es_row_no-row_id INTO ls_row_data.

          IF sy-subrc = 0 AND ls_row_data-belnr IS NOT INITIAL.

            SET PARAMETER ID 'BLN' FIELD ls_row_data-belnr.
            SET PARAMETER ID 'BUK' FIELD ls_row_data-bukrs.
            SET PARAMETER ID 'GJR' FIELD ls_row_data-gjahr.

            CALL TRANSACTION 'FB03' AND SKIP FIRST SCREEN.
          ENDIF.
        ENDIF.
       when 'LIFNR'.
         IF es_row_no-row_id is not initial.
           assign lr_data->* to <lt_table>.
           read table <lt_table> index es_row_no-row_id into ls_row_data.

           IF sy-subrc eq 0.
             set PARAMETER ID 'BUK' FIELD ls_row_data-bukrs.
             SET PARAMETER ID 'LIF' FIELD ls_row_data-lifnr.
             call TRANSACTION 'FBL1N' and SKIP FIRST SCREEN.
           ENDIF.
         ENDIF.
    ENDCASE.
  ENDMETHOD.

  METHOD on_button_click. ENDMETHOD.

METHOD on_top_of_page.
    CALL METHOD e_dyndoc_id->add_picture EXPORTING picture_id = 'FATURALAB'.
ENDMETHOD.

ENDCLASS.
CLASS lcl_alv IMPLEMENTATION.
  METHOD constructor.
    mv_tabname = i_str.
    CREATE DATA mt_outtab TYPE STANDARD TABLE OF (mv_tabname) WITH EMPTY KEY.
  ENDMETHOD.


  METHOD create.
    DATA(lo_alv) = NEW lcl_alv( i_str = i_str ).
    DATA: lo_cont TYPE REF TO cl_gui_custom_container.
     DATA go_logo TYPE REF TO cl_gui_picture.
    CREATE OBJECT lo_cont EXPORTING container_name = 'CUSTOM'.
     DATA(lo_splitter) = NEW cl_gui_splitter_container( parent = lo_cont rows = 2 columns = 1 ).
      DATA(lo_cont_logo) = lo_splitter->get_container( row = 1 column = 1 ).
      data(lo_cont_alv) = lo_splitter->get_container( row = 2 column = 1 ).
      lo_splitter->set_row_height( id = 1 height = 15 ).

     CREATE OBJECT go_logo
      EXPORTING
        parent = lo_cont_logo.
    DATA(lv_url) = lo_alv->get_logo_url( ).

    CALL METHOD go_logo->load_picture_from_url
      EXPORTING
        url = lv_url.
    CREATE OBJECT lo_alv->go_grd EXPORTING i_parent = lo_cont_alv.
    lo_alv->get_data( ).
    DATA(lo_eh) = NEW lcl_eh_head( i_grid = lo_alv->go_grd i_rpt = lo_alv ).
    lo_alv->cre_fct( ).
    lo_alv->cre_lay( ).
    e_alv = lo_alv.
  ENDMETHOD.
METHOD get_data.
  FIELD-SYMBOLS <lt_data> TYPE INDEX TABLE.
  DATA: lt_style TYPE lvc_t_styl,
        ls_style TYPE lvc_s_styl.

  ASSIGN mt_outtab->* TO <lt_data>.

  SELECT * FROM (mv_tabname)
    INTO CORRESPONDING FIELDS OF TABLE @<lt_data>
    ORDER BY belnr DESCENDING.

  LOOP AT <lt_data> ASSIGNING FIELD-SYMBOL(<ls_row>).
    CLEAR lt_style.

    ASSIGN COMPONENT 'BUKRS' OF STRUCTURE <ls_row> TO FIELD-SYMBOL(<lv_bukrs>).
    ASSIGN COMPONENT 'LIFNR' OF STRUCTURE <ls_row> to FIELD-SYMBOL(<lv_lifnr>).
    IF <lv_bukrs> IS ASSIGNED.
      AUTHORITY-CHECK OBJECT 'F_BKPF_BUK'
        ID 'BUKRS' FIELD <lv_bukrs>
        ID 'ACTVT' FIELD '02'.

      IF sy-subrc <> 0.
        ls_style-style = cl_gui_alv_grid=>mc_style_disabled.
        ls_style-fieldname = 'BUKRS'. APPEND ls_style TO lt_style.
        ls_style-fieldname = 'AMO'.   APPEND ls_style TO lt_style.
        ls_style-fieldname = 'WAERS'. APPEND ls_style TO lt_style.
      ENDIF.
    ENDIF.

    ASSIGN COMPONENT 'STYLE' OF STRUCTURE <ls_row> TO FIELD-SYMBOL(<lt_row_style>).
    IF <lt_row_style> IS ASSIGNED.
      <lt_row_style> = lt_style.
    ENDIF.
  ENDLOOP.
ENDMETHOD.

  METHOD cre_fct.
    CALL FUNCTION 'LVC_FIELDCATALOG_MERGE'
      EXPORTING i_structure_name = mv_tabname
      CHANGING ct_fieldcat = gt_fct
      EXCEPTIONS OTHERS = 3.
    me->set_alv_hot( 'BELNR' ).
    me->set_alv_hot( 'LIFNR' ).
    LOOP AT gt_fct REFERENCE INTO DATA(lr_fct).
    CASE lr_fct->fieldname.
      WHEN 'BUKRS' OR 'BLART' OR 'MONAT' OR 'WAERS' OR 'AMO' OR 'BLDAT'.
        lr_fct->edit = abap_true.
      WHEN 'BELNR'.
        lr_fct->hotspot = abap_true.
      when 'GJAHR'.
        lr_fct->col_opt = abap_true.
    ENDCASE.
  ENDLOOP.

  ENDMETHOD.
  METHOD cre_lay.
  gs_lay = VALUE #(
    zebra      = 'X'
    sel_mode   = 'A'
    cwidth_opt = 'X'
    stylefname = 'STYLE'
  ).
  ENDMETHOD.

  METHOD dis_alv.
    data: lt_excluding type ui_functions.
    APPEND cl_gui_alv_grid=>mc_fc_loc_insert_row to lt_excluding.
    append cl_gui_alv_grid=>mc_fc_loc_delete_row to lt_excluding.
    APPEND cl_gui_alv_grid=>mc_fc_help to lt_excluding.
    FIELD-SYMBOLS <lt_data> TYPE INDEX TABLE.
    ASSIGN mt_outtab->* TO <lt_data>.
    go_grd->set_table_for_first_display(
      EXPORTING is_layout = gs_lay
                it_toolbar_excluding = lt_excluding
      CHANGING it_outtab = <lt_data> it_fieldcatalog = gt_fct ).
  ENDMETHOD.

  METHOD set_alv_hot.
    READ TABLE gt_fct REFERENCE INTO DATA(lr_fct) WITH KEY fieldname = i_fnam.
    IF sy-subrc = 0. lr_fct->hotspot = abap_true. ENDIF.
  ENDMETHOD.

 METHOD disp_mess_popup.
  CALL METHOD cl_reca_gui_f4_popup=>factory_grid
    EXPORTING
      it_f4value   = it_mess[]
      id_title     = 'İşlem Sırasında Alınan Hatalar'
    RECEIVING
      ro_f4_instance = DATA(lo_popup).

  lo_popup->display( id_start_column = 25 id_start_line = 10 ).
ENDMETHOD.
  METHOD get_logo_url.
    CLEAR r_url.

    DATA query_table    TYPE STANDARD TABLE OF w3query.
    DATA ls_query_table TYPE w3query.
    DATA html_table     TYPE STANDARD TABLE OF w3html .
    DATA return_code    TYPE  w3param-ret_code.
    DATA content_type   TYPE  w3param-cont_type.
    DATA content_length TYPE  w3param-cont_len.
    DATA pic_data       TYPE STANDARD TABLE OF w3mime.
    DATA pic_size       TYPE i.

    REFRESH query_table.
    ls_query_table-name = '_OBJECT_ID'.
    ls_query_table-value = 'YFATURALAB_LOGO'.
    APPEND ls_query_table TO query_table.

    CALL FUNCTION 'WWW_GET_MIME_OBJECT'
      TABLES
        query_string        = query_table
        html                = html_table
        mime                = pic_data
      CHANGING
        return_code         = return_code
        content_type        = content_type
        content_length      = content_length
      EXCEPTIONS
        object_not_found    = 1
        parameter_not_found = 2
        OTHERS              = 3.
    IF sy-subrc = 0.
      pic_size = content_length.
    ENDIF.

    CALL FUNCTION 'DP_CREATE_URL'
      EXPORTING
        type     = 'image'
        subtype  = cndp_sap_tab_unknown
        size     = pic_size
        lifetime = cndp_lifetime_transaction
      TABLES
        data     = pic_data
      CHANGING
        url      = r_url
      EXCEPTIONS
        OTHERS   = 1.
  ENDMETHOD.
ENDCLASS.
FORM z_excel_vendor.
  DATA: lv_belnr         TYPE belnr_d,
        ls_zfat          TYPE zfat,
        lt_vendor_ex     TYPE TABLE OF ty_excel,
        ls_vendor_ex     TYPE ty_excel,
        lt_all_errors    TYPE lcl_alv=>ty_mess_pop_t,
        lv_msg_text(200) TYPE c.

  DATA: lv_bukrs_bdc TYPE bdc_fval,
        lv_blart_bdc TYPE bdc_fval,
        lv_dat_bdc   TYPE bdc_fval,
        lv_monat_bdc TYPE bdc_fval,
        lv_waers_bdc TYPE bdc_fval,
        lv_amo_bdc   TYPE bdc_fval,
        lv_newko_bdc TYPE bdc_fval,
        lv_newk2_bdc TYPE bdc_fval.

  CALL FUNCTION 'TEXT_CONVERT_XLS_TO_SAP'
    EXPORTING
      i_line_header        = 'X'
      i_filename           = gv_file2
      i_tab_raw_data       = lt_raw
    TABLES
      i_tab_converted_data = lt_vendor_ex
    EXCEPTIONS
      conversion_failed    = 1
      OTHERS               = 2.

  IF sy-subrc <> 0.
    MESSAGE 'Excel verileriniz okunamadı' TYPE 'E'.
    RETURN.
  ENDIF.


  LOOP AT lt_vendor_ex INTO ls_vendor_ex.
    CLEAR: lt_messtab, lv_subrc, lv_dat_bdc, lv_amo_bdc, lv_newko_bdc, lv_newk2_bdc.

    IF ls_vendor_ex-dat IS NOT INITIAL.
      CALL FUNCTION 'CONVERT_DATE_TO_EXTERNAL'
        EXPORTING
          date_internal            = ls_vendor_ex-dat
        IMPORTING
          date_external            = lv_dat_bdc
        EXCEPTIONS
          error_internal_date_type = 1
          OTHERS                   = 2.
      IF sy-subrc <> 0.
        WRITE ls_vendor_ex-dat TO lv_dat_bdc.
      ENDIF.
    ENDIF.

    WRITE ls_vendor_ex-amo TO lv_amo_bdc.
    CONDENSE lv_amo_bdc NO-GAPS.

    lv_bukrs_bdc = ls_vendor_ex-bukrs.
    lv_blart_bdc = ls_vendor_ex-blart.
    lv_monat_bdc = ls_vendor_ex-monat.
    lv_waers_bdc = ls_vendor_ex-waers.
    lv_newko_bdc = ls_vendor_ex-newko.
    lv_newk2_bdc = ls_vendor_ex-newk2.

    CALL FUNCTION 'ZF2_VENDOR2'
      EXPORTING
        ctu       = 'X'
        mode      = 'N'
        update    = 'S'
        nodata    = '/'
        bldat_001 = lv_dat_bdc
        blart_002 = lv_blart_bdc
        bukrs_003 = lv_bukrs_bdc
        budat_004 = lv_dat_bdc
        monat_005 = lv_monat_bdc
        waers_006 = lv_waers_bdc
        newko_009 = lv_newko_bdc
        wrbtr_010 = lv_amo_bdc
        newko_017 = lv_newk2_bdc
        wrbtr_013 = lv_amo_bdc
        wrbtr_018 = lv_amo_bdc
        wrbtr_021 = lv_amo_bdc
        zfbdt_012 = lv_dat_bdc
        zfbdt_015 = lv_dat_bdc
        zfbdt_020 = lv_dat_bdc
        zfbdt_023 = lv_dat_bdc
      IMPORTING
        subrc     = lv_subrc
      TABLES
        messtab   = lt_messtab.

    IF lv_subrc = 0.
      READ TABLE lt_messtab INTO DATA(ls_msg) WITH KEY msgid = 'F5' msgnr = '312'.
      IF sy-subrc = 0.
        lv_belnr = ls_msg-msgv1.
        ls_zfat = VALUE #(
          mandt = sy-mandt
          bukrs = ls_vendor_ex-bukrs
          belnr = lv_belnr
          gjahr = ls_vendor_ex-dat(4)
          blart = ls_vendor_ex-blart
          monat = ls_vendor_ex-monat
          waers = ls_vendor_ex-waers
          amo   = ls_vendor_ex-amo
          uname = sy-uname
          lifnr = ls_vendor_ex-newko
          cpudt = sy-datum
          cputm = sy-uzeit
        ).
        MODIFY zfat FROM ls_zfat.
        COMMIT WORK AND WAIT.
      ENDIF.
    ELSE.
      LOOP AT lt_messtab INTO DATA(ls_bdc_msg) WHERE msgtyp = 'E' OR msgtyp = 'A'.
        CLEAR lv_msg_text.
        MESSAGE ID     ls_bdc_msg-msgid
                TYPE   ls_bdc_msg-msgtyp
                NUMBER ls_bdc_msg-msgnr
                WITH   ls_bdc_msg-msgv1 ls_bdc_msg-msgv2 ls_bdc_msg-msgv3 ls_bdc_msg-msgv4
                INTO   lv_msg_text.

        APPEND VALUE #( statu = icon_led_red hata = lv_msg_text ) TO lt_all_errors.
      ENDLOOP.
    ENDIF.
  ENDLOOP.

  IF lt_all_errors IS NOT INITIAL.
    IF go_alv IS NOT BOUND.
      go_alv = lcl_alv=>create( i_str = 'ZFAT' ).
    ENDIF.
    go_alv->disp_mess_popup( it_mess = lt_all_errors ).
  ENDIF.

  IF go_alv IS NOT BOUND.
    go_alv = lcl_alv=>create( i_str = 'ZFAT' ).
  ENDIF.

  go_alv->get_data( ).
  CALL SCREEN 103.

ENDFORM.
FORM get_balance_data.
  SELECT b~lifnr, l~name1, b~shkzg, b~wrbtr, b~waers
    FROM bsik AS b
    INNER JOIN lfa1 AS l ON b~lifnr = l~lifnr
    INTO TABLE @DATA(lt_items)
    WHERE b~lifnr IN @s_lifnr.

  FIELD-SYMBOLS: <fs_balance> LIKE LINE OF lt_balance.
  DATA: ls_color TYPE lvc_s_scol.

  LOOP AT lt_items INTO DATA(ls_item).
    READ TABLE lt_balance ASSIGNING <fs_balance> WITH KEY lifnr = ls_item-lifnr.
    IF sy-subrc <> 0.
      APPEND VALUE #( lifnr = ls_item-lifnr
                      name1 = ls_item-name1
                      waers = ls_item-waers ) TO lt_balance ASSIGNING <fs_balance>.
    ENDIF.

    IF ls_item-shkzg = 'H'.
      <fs_balance>-t_girdi = <fs_balance>-t_girdi + ls_item-wrbtr.
    ELSEIF ls_item-shkzg = 'S'.
      <fs_balance>-t_cikti = <fs_balance>-t_cikti + ls_item-wrbtr.
    ENDIF.
  ENDLOOP.

  LOOP AT lt_balance ASSIGNING <fs_balance>.
    <fs_balance>-bakiye = <fs_balance>-t_girdi - <fs_balance>-t_cikti.

    REFRESH <fs_balance>-t_color.
    ls_color-fname = 'BAKIYE'.

    IF <fs_balance>-bakiye > 0.
      ls_color-color-col = 5.
    ELSEIF <fs_balance>-bakiye = 0.
      ls_color-color-col = 1.
    ELSE.
      ls_color-color-col = 6.
    ENDIF.
    APPEND ls_color TO <fs_balance>-t_color.
  ENDLOOP.
ENDFORM.
FORM display_balance_report.
  PERFORM get_balance_data.

  DATA: lt_fcat TYPE lvc_t_fcat,
        ls_fcat TYPE lvc_s_fcat,
        ls_layo TYPE lvc_s_layo.

  ls_layo-ctab_fname = 'T_COLOR'.
  ls_layo-cwidth_opt = 'X'.

  DEFINE m_append_fcat.
    CLEAR ls_fcat.
    ls_fcat-fieldname = &1.
    ls_fcat-scrtext_m = &2.
    ls_fcat-do_sum    = &3.
    APPEND ls_fcat TO lt_fcat.
  END-OF-DEFINITION.

  m_append_fcat 'LIFNR'   'Satıcı No'     ' '.
  m_append_fcat 'NAME1'   'Satıcı Adı'    ' '.
  m_append_fcat 'T_GIRDI' 'Toplam Alacak' 'X'.
  m_append_fcat 'T_CIKTI' 'Toplam Borç'   'X'.
  m_append_fcat 'BAKIYE'  'Net Bakiye'    'X'.
  m_append_fcat 'WAERS'   'Para Birimi'   ' '.

  CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY_LVC'
    EXPORTING
      is_layout_lvc   = ls_layo
      it_fieldcat_lvc = lt_fcat
      i_grid_title    = 'Satıcı Bakiye Analiz Raporu'
    TABLES
      t_outtab        = lt_balance
    EXCEPTIONS
      program_error   = 1
      OTHERS          = 2.
ENDFORM.
MODULE status_0100 OUTPUT.
  SET PF-STATUS '100'.
  IF rb2 = 'X'.
    gv_subscreen = '0101'.
  ELSE.
    gv_subscreen = '0102'.
  ENDIF.

  LOOP AT SCREEN.
    IF screen-group1 = 'GR1'.
      IF rb1 = 'X'. screen-active = '1'. ELSE. screen-active = '0'. ENDIF.
      MODIFY SCREEN.
    ENDIF.
  ENDLOOP.

ENDMODULE.
MODULE user_command_0100 INPUT.
  CASE sy-ucomm.
    WHEN '&EXC'.
      CALL FUNCTION 'F4_FILENAME' IMPORTING file_name = gv_file.
      IF gv_file IS NOT INITIAL. PERFORM z_excel. ENDIF.

    WHEN '&EXC2'.
      CALL FUNCTION 'F4_FILENAME' IMPORTING file_name = gv_file2.
      IF gv_file2 IS NOT INITIAL. PERFORM z_excel_vendor. ENDIF.

 WHEN 'SAVE'.
  DATA: lv_dat_str   TYPE bdc_fval,
        lv_amo_str   TYPE bdc_fval,
        lv_bukrs_str TYPE bdc_fval,
        lv_monat_str TYPE bdc_fval,
        lv_waers_str TYPE bdc_fval,
        lv_newko_str TYPE bdc_fval,
        lv_newk2_str TYPE bdc_fval,
        lv_blart_str TYPE bdc_fval,
        lv_belnr     TYPE belnr_d,
        ls_zfat      TYPE zfat,
        lt_all_errors TYPE lcl_alv=>ty_mess_pop_t,
        lv_msg_text(200) TYPE c.

  IF waers IS INITIAL OR amo IS INITIAL OR bldat IS INITIAL.
    MESSAGE 'Lütfen tarih, para birimi ve tutarı giriniz!' TYPE 'S' DISPLAY LIKE 'E'.
    RETURN.
  ENDIF.

  CALL FUNCTION 'CONVERT_DATE_TO_EXTERNAL'
    EXPORTING
      date_internal = bldat
    IMPORTING
      date_external = lv_dat_str.

  lv_bukrs_str = bukrs.
  lv_blart_str = blart.
  lv_waers_str = waers.
  lv_monat_str = monat.
  lv_newko_str = newko.
  lv_newk2_str = newk2.

  WRITE amo TO lv_amo_str CURRENCY waers.
  CONDENSE lv_amo_str NO-GAPS.

  CLEAR lt_messtab.
  CALL FUNCTION 'ZF_02'
    EXPORTING
      ctu       = 'X'
      mode      = 'N'
      update    = 'S'
      bldat_001 = lv_dat_str
      blart_002 = lv_blart_str
      bukrs_003 = lv_bukrs_str
      budat_004 = lv_dat_str
      monat_005 = lv_monat_str
      waers_006 = lv_waers_str
      wwert_007 = lv_dat_str
      newko_010 = lv_newko_str
      wrbtr_011 = lv_amo_str
      newko_013 = lv_newk2_str
      wrbtr_015 = lv_amo_str
      wrbtr_018 = lv_amo_str
    IMPORTING
      subrc     = lv_subrc
    TABLES
      messtab   = lt_messtab.

  IF lv_subrc = 0.
    READ TABLE lt_messtab INTO DATA(ls_msg) WITH KEY msgid = 'F5' msgnr = '312'.
    IF sy-subrc = 0.
      lv_belnr = ls_msg-msgv1.

      IF lv_belnr IS NOT INITIAL.
        CLEAR ls_zfat.
        ls_zfat = VALUE #(
          mandt = sy-mandt
          bukrs = lv_bukrs_str
          belnr = lv_belnr
          gjahr = bldat(4)
          blart = lv_blart_str
          monat = lv_monat_str
          waers = lv_waers_str
          amo   = amo
          uname = sy-uname
          cpudt = sy-datum
          lifnr = lv_newko_str
          cputm = sy-uzeit
        ).

        MODIFY zfat FROM ls_zfat.
        COMMIT WORK AND WAIT.

        MESSAGE 'Kayıt başarıyla oluşturuldu: ' && lv_belnr TYPE 'S'.

        IF go_alv IS NOT BOUND.
          go_alv = lcl_alv=>create( i_str = 'ZFAT' ).
        ELSE.
          go_alv->get_data( ).
        ENDIF.
        CALL SCREEN 103.
      ENDIF.
    ENDIF.
  ELSE.
    LOOP AT lt_messtab INTO DATA(ls_bdc_msg) WHERE msgtyp = 'E' OR msgtyp = 'A'.
      CLEAR lv_msg_text.
      MESSAGE ID     ls_bdc_msg-msgid
              TYPE   ls_bdc_msg-msgtyp
              NUMBER ls_bdc_msg-msgnr
              WITH   ls_bdc_msg-msgv1 ls_bdc_msg-msgv2 ls_bdc_msg-msgv3 ls_bdc_msg-msgv4
              INTO   lv_msg_text.

      APPEND VALUE #( statu = icon_led_red hata = lv_msg_text ) TO lt_all_errors.
    ENDLOOP.

    IF lt_all_errors IS NOT INITIAL.
      IF go_alv IS NOT BOUND.
        go_alv = lcl_alv=>create( i_str = 'ZFAT' ).
      ENDIF.
      go_alv->disp_mess_popup( it_mess = lt_all_errors ).
    ELSE.
      MESSAGE 'Kayıt sırasında hata oluştu, ancak mesaj dönmedi!' TYPE 'W'.
    ENDIF.
  ENDIF.

    WHEN '&EXIT' OR '&BACK' OR '&CANCEL'.
      LEAVE TO SCREEN 0.
  ENDCASE.
ENDMODULE.
MODULE f4_date_help INPUT.
  CALL FUNCTION 'F4_DATE'
    EXPORTING date_for_first_month = sy-datum
    IMPORTING select_date          = bldat.

ENDMODULE.
MODULE f4_bukrs_help INPUT.
CALL FUNCTION 'F4IF_FIELD_VALUE_REQUEST'
  EXPORTING
    tabname = 'ZFAT'
    fieldname = 'BUKRS'
    dynpprog = sy-repid
    dynpnr = sy-dynnr
    dynprofield = 'BUKRS'
    EXCEPTIONS
      field_not_found   = 1
      no_help_for_field = 2
      inconsistent_help = 3
      no_values_found   = 4
      others            = 5.

ENDMODULE.
MODULE f4_blart_help INPUT.
  CALL FUNCTION 'F4IF_FIELD_VALUE_REQUEST'
    EXPORTING
      tabname           = 'BKPF'
      fieldname         = 'BLART'
      dynpprog          = sy-repid
      dynpnr            = sy-dynnr
      dynprofield       = 'BLART'
    EXCEPTIONS
      others            = 1.
ENDMODULE.
MODULE f4_waers_help INPUT.
  CALL FUNCTION 'F4IF_FIELD_VALUE_REQUEST'
    EXPORTING
      tabname           = 'BKPF'
      fieldname         = 'WAERS'
      dynpprog          = sy-repid
      dynpnr            = sy-dynnr
      dynprofield       = 'WAERS'
    EXCEPTIONS
      others            = 1.
ENDMODULE.
FORM set_ini.
call method yflb_cl_faturalab=>get_user_comp_from_authority
  EXPORTING
    i_uname      =  sy-uname
  RECEIVING
    e_bukrs      =  data(lv_bukrs)
  EXCEPTIONS
    no_authority = 1
    others       = 2
  .
IF SY-SUBRC <> 0.
message i014(yflb).
leave program.
ENDIF.
ENDFORM.
FORM bukrs .
  DATA: lt_ust12 TYPE TABLE OF ust12 WITH HEADER LINE.
data: gt_bukrs type vrm_values.
data: gwa_bukrs type vrm_value.
data: name type vrm_id.
  REFRESH gt_bukrs.
  CLEAR: gwa_bukrs, gt_bukrs[].
  name = 'p_bukrs'.
  SELECT * FROM ust12 INTO TABLE lt_ust12 WHERE objct = 'YFLB_COMP'.

  LOOP AT lt_ust12 WHERE von NE '*' .
    gwa_bukrs-key  = lt_ust12-von.
    gwa_bukrs-text = lt_ust12-von.
    COLLECT gwa_bukrs INTO gt_bukrs.
  ENDLOOP.
  CALL FUNCTION 'VRM_SET_VALUES'
    EXPORTING
      id              = name
      values          = gt_bukrs
    EXCEPTIONS
      id_illegal_name = 1
      OTHERS          = 2.
ENDFORM.
*&---------------------------------------------------------------------*
*& Module STATUS_0103 OUTPUT
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
MODULE status_0103 OUTPUT.
  SET PF-STATUS '103'.

  FIELD-SYMBOLS <lt_outtab> TYPE INDEX TABLE.

  IF go_alv IS BOUND.
    go_alv->dis_alv( ).
  ENDIF.

ENDMODULE.
*&---------------------------------------------------------------------*
*&      Module  USER_COMMAND_0103  INPUT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
MODULE user_command_0103 INPUT.
  CASE sy-ucomm.
    WHEN 'VEND'.
    PERFORM display_balance_report.

    WHEN '&EXIT' OR '&BACK' OR '&CANCEL'.
      LEAVE TO SCREEN 100.
  ENDCASE.
ENDMODULE.
*&---------------------------------------------------------------------*
*& Module STATUS_0104 OUTPUT
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
MODULE status_0104 OUTPUT.
  SET PF-STATUS '104'.

ENDMODULE.
DATA: gt_items TYPE TABLE OF rfposxext.
FORM display_balance.
  CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
    EXPORTING
      i_structure_name = 'RFPOSXEXT'
      i_grid_title     = 'Satıcı Bakiye Detayları'
    TABLES
      t_outtab         = gt_items
    EXCEPTIONS
      program_error    = 1
      OTHERS           = 2.
ENDFORM.
*&---------------------------------------------------------------------*
*&      Module  USER_COMMAND_0104  INPUT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
MODULE user_command_0104 INPUT.
CASE sy-ucomm.
  WHEN '&EXIT' OR '&CANCEL' OR '&BACK'.
    LEAVE TO SCREEN 103.
ENDCASE.
ENDMODULE.