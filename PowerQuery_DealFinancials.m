// =============================================================================
// Query: _DealFinancials
// Pegar en: Excel > Datos > Obtener datos > De otras fuentes > Consulta en blanco
//           (Data > Get Data > From Other Sources > Blank Query) > Editor avanzado
// Luego: Cerrar y cargar EN > Hoja nueva (o en la hoja "_DealFinancialsData")
//        y ocultar la hoja.
//
// Replica la misma lógica que el PBI (Deal Financials + UW Financials unidas),
// devuelve una tabla única con la columna [Category] = "ACT" o "UW".
// =============================================================================
let
    Origen                  = CommonDataService.Database("org98d6617c.crm.dynamics.com"),
    dbo_assetmanagement     = Origen{[Schema="dbo", Item="cr83c_assetmanagement"]}[Data],
    dbo_loancollateral      = Origen{[Schema="dbo", Item="cr83c_loancollateral"]}[Data],

    // ------------- Base común -------------
    Active                  = Table.SelectRows(dbo_assetmanagement, each [cr83c_isactive] = true),

    JoinCollateral          = Table.NestedJoin(
                                Active, {"cr83c_loancollateral"},
                                dbo_loancollateral, {"cr83c_loancollateralid"},
                                "LoanCollateral_Record", JoinKind.LeftOuter),

    ExpandCollateral        = Table.ExpandTableColumn(
                                JoinCollateral, "LoanCollateral_Record",
                                {"cr83c_propertysize"}, {"Property Size"}),

    // ------------- Diccionario (para rescatar nombres nulos en UW) -------------
    DictExpand              = Table.ExpandTableColumn(
                                Active, "cr83c_dealfinancials",
                                {
                                    "cr83c_accountcategory", "cr83c_accountcategoryname",
                                    "cr83c_accountsubcategory", "cr83c_accountsubcategoryname",
                                    "cr83c_internallineitemname"
                                },
                                {
                                    "Account Category ID", "Account Category",
                                    "Sub-Category ID", "Sub-Category",
                                    "Internal Line Item"
                                }),
    DictCols                = Table.SelectColumns(DictExpand,
                                {"Account Category ID", "Account Category",
                                 "Sub-Category ID", "Sub-Category", "Internal Line Item"}),
    DictNotNull             = Table.SelectRows(DictCols, each [Account Category ID] <> null),
    DictTyped               = Table.TransformColumnTypes(DictNotNull, {
                                {"Account Category", type text}, {"Sub-Category", type text},
                                {"Internal Line Item", type text},
                                {"Account Category ID", type text}, {"Sub-Category ID", type text}
                              }),
    DictNormIds             = Table.TransformColumns(DictTyped, {
                                {"Account Category ID", each try Text.Lower(Text.From(_)) otherwise null, type text},
                                {"Sub-Category ID",     each try Text.Lower(Text.From(_)) otherwise null, type text}
                              }),
    // --- Aplicamos el mismo fix que el PBI hace en 'UID - Categories' antes de deduplicar ---
    DictFix                 = Table.AddColumn(DictNormIds, "FixedData", each
                                let
                                    AcID   = [#"Account Category ID"],
                                    ScID   = [#"Sub-Category ID"],
                                    CurrAc = [#"Account Category"],
                                    CurrSc = [#"Sub-Category"]
                                in
                                    if ScID = "591900014" and AcID = "591900007" then [AC="Below the line", SC="CapEx Reserve"]
                                    else if ScID = "591900017" and AcID = "591900000" then [AC=CurrAc, SC="Other Income"]
                                    else if ScID = "591900015" and AcID = "591900001" then [AC=CurrAc, SC="HOA"]
                                    else if ScID = "591900016" and AcID = "591900002" then [AC=CurrAc, SC="Ground Lease"]
                                    else [AC=CurrAc, SC=CurrSc]
                              ),
    DictFixExpand           = Table.ExpandRecordColumn(DictFix, "FixedData", {"AC", "SC"},
                                {"Account Category Fixed", "Sub-Category Fixed"}),
    // Mismo patrón que 'UID - Categories' del PBI: primero seleccionamos las "Fixed"
    // (descartando las originales), y después renombramos a los nombres finales.
    DictPick                = Table.SelectColumns(DictFixExpand,
                                {"Account Category ID", "Account Category Fixed",
                                 "Sub-Category ID", "Sub-Category Fixed", "Internal Line Item"}),
    DictFinal               = Table.RenameColumns(DictPick, {
                                {"Account Category Fixed", "Account Category"},
                                {"Sub-Category Fixed",     "Sub-Category"}
                              }),
    DictUnique              = Table.Distinct(DictFinal, {"Internal Line Item"}),

    // ------------- Función reutilizable: corrige AC/SC según IDs -------------
    FixCategories           = (origAC as nullable text, origSC as nullable text,
                               acID as nullable text,  scID as nullable text,
                               rescueAC as nullable text, rescueSC as nullable text) =>
        if scID = "591900014" and acID = "591900007" then [AC="Below the line", SC="CapEx Reserve"]
        else if scID = "591900017" and acID = "591900000" then [AC = (if origAC=null then rescueAC else origAC), SC="Other Income"]
        else if scID = "591900015" and acID = "591900001" then [AC = (if origAC=null then rescueAC else origAC), SC="HOA"]
        else if scID = "591900016" and acID = "591900002" then [AC = (if origAC=null then rescueAC else origAC), SC="Ground Lease"]
        else [
            AC = if origAC <> null then origAC else if rescueAC <> null then rescueAC else "Unassigned",
            SC = if origSC <> null then origSC else if rescueSC <> null then rescueSC else "Unassigned"
        ],

    // ------------- ACT: cr83c_dealfinancials -------------
    ActExpand               = Table.ExpandTableColumn(
                                ExpandCollateral, "cr83c_dealfinancials",
                                {
                                    "cr83c_date", "cr83c_amount",
                                    "cr83c_accountcategoryname", "cr83c_accountcategory",
                                    "cr83c_accountsubcategoryname", "cr83c_accountsubcategory",
                                    "cr83c_comments", "cr83c_internallineitemname"
                                },
                                {
                                    "Date", "Amount",
                                    "Account Category Name", "Account Category ID",
                                    "Sub-Category Name", "Sub-Category ID",
                                    "Comments", "Internal Line Item"
                                }),
    ActNormIds              = Table.TransformColumns(ActExpand, {
                                {"Account Category ID", each try Text.Lower(Text.From(_)) otherwise null, type text},
                                {"Sub-Category ID",     each try Text.Lower(Text.From(_)) otherwise null, type text}
                              }),
    ActFix                  = Table.AddColumn(ActNormIds, "FixedData", each
                                FixCategories(
                                    [#"Account Category Name"], [#"Sub-Category Name"],
                                    [#"Account Category ID"],   [#"Sub-Category ID"],
                                    null, null
                                )),
    ActFixExpand            = Table.ExpandRecordColumn(ActFix, "FixedData", {"AC", "SC"},
                                {"Account Category", "Sub-Category"}),
    ActSel                  = Table.SelectColumns(ActFixExpand,
                                {"Date", "cr83c_name", "cr83c_dealnumber", "Property Size",
                                 "Amount", "Account Category", "Sub-Category",
                                 "Internal Line Item", "Comments", "cr83c_originalcommitment"}),
    ActRen                  = Table.RenameColumns(ActSel, {
                                {"cr83c_originalcommitment", "Loan Amount"},
                                {"cr83c_name", "Deal Name"},
                                {"cr83c_dealnumber", "Deal Number"}
                              }),
    ActTyped                = Table.TransformColumnTypes(ActRen, {
                                {"Date", type date}, {"Amount", Currency.Type},
                                {"Property Size", Int64.Type}, {"Loan Amount", Currency.Type},
                                {"Account Category", type text}, {"Sub-Category", type text}
                              }),
    ActFinal                = Table.AddColumn(ActTyped, "Category", each "ACT", type text),

    // ------------- UW: cr83c_underwritingfinancials -------------
    // NOTA: Si tu conector Dataverse de Excel no expone este nombre de relación
    // cambiá "cr83c_underwritingfinancials" por el que corresponda
    // (podés listarlos con: Table.ColumnNames( dbo_assetmanagement ) ).
    UwRelationColumnName    = "cr83c_underwritingfinancials",
    UwColumnExists          = List.Contains(Table.ColumnNames(ExpandCollateral), UwRelationColumnName),
    UwExpand                = if UwColumnExists then
                                Table.ExpandTableColumn(
                                    ExpandCollateral, UwRelationColumnName,
                                    {
                                        "cr83c_date", "cr83c_amount",
                                        "cr83c_accountcategoryname", "cr83c_accountcategory",
                                        "cr83c_accountsubcategoryname", "cr83c_accountsubcategory",
                                        "cr83c_comments", "cr83c_internallineitemname"
                                    },
                                    {
                                        "Date", "Amount",
                                        "Account Category Name", "Account Category ID",
                                        "Sub-Category Name", "Sub-Category ID",
                                        "Comments", "Internal Line Item"
                                    })
                              else
                                // Fallback: tabla vacía con el schema esperado para que el combine siga andando
                                #table(
                                    type table [
                                        #"cr83c_name" = text, #"cr83c_dealnumber" = text,
                                        #"Property Size" = Int64.Type,
                                        #"cr83c_originalcommitment" = Currency.Type,
                                        Date = date, Amount = Currency.Type,
                                        #"Account Category Name" = text, #"Account Category ID" = text,
                                        #"Sub-Category Name" = text, #"Sub-Category ID" = text,
                                        Comments = text, #"Internal Line Item" = text
                                    ],
                                    {}
                                ),
    UwNormIds               = Table.TransformColumns(UwExpand, {
                                {"Account Category ID", each try Text.Lower(Text.From(_)) otherwise null, type text},
                                {"Sub-Category ID",     each try Text.Lower(Text.From(_)) otherwise null, type text}
                              }),
    UwJoinDict              = Table.NestedJoin(UwNormIds, {"Internal Line Item"},
                                DictUnique, {"Internal Line Item"},
                                "Diccionario_Record", JoinKind.LeftOuter),
    UwExpandDict            = Table.ExpandTableColumn(UwJoinDict, "Diccionario_Record",
                                {"Account Category", "Sub-Category"},
                                {"M_AcName", "M_ScName"}),
    UwFix                   = Table.AddColumn(UwExpandDict, "FixedData", each
                                FixCategories(
                                    [#"Account Category Name"], [#"Sub-Category Name"],
                                    [#"Account Category ID"],   [#"Sub-Category ID"],
                                    [M_AcName], [M_ScName]
                                )),
    UwFixExpand             = Table.ExpandRecordColumn(UwFix, "FixedData", {"AC", "SC"},
                                {"Account Category", "Sub-Category"}),
    UwSel                   = Table.SelectColumns(UwFixExpand,
                                {"Date", "cr83c_name", "cr83c_dealnumber", "Property Size",
                                 "Amount", "Account Category", "Sub-Category",
                                 "Internal Line Item", "Comments", "cr83c_originalcommitment"}),
    UwRen                   = Table.RenameColumns(UwSel, {
                                {"cr83c_originalcommitment", "Loan Amount"},
                                {"cr83c_name", "Deal Name"},
                                {"cr83c_dealnumber", "Deal Number"}
                              }),
    UwTyped                 = Table.TransformColumnTypes(UwRen, {
                                {"Date", type date}, {"Amount", Currency.Type},
                                {"Property Size", Int64.Type}, {"Loan Amount", Currency.Type},
                                {"Account Category", type text}, {"Sub-Category", type text}
                              }),
    UwFinal                 = Table.AddColumn(UwTyped, "Category", each "UW", type text),

    // ------------- Unión -------------
    Combined                = Table.Combine({ActFinal, UwFinal}),

    // Orden de columnas final (lo usarán los SUMIFS del Excel)
    Reordered               = Table.ReorderColumns(Combined, {
                                "Deal Name", "Deal Number", "Date", "Property Size",
                                "Category", "Account Category", "Sub-Category",
                                "Internal Line Item", "Amount", "Comments", "Loan Amount"
                              })
in
    Reordered
