-- EC ISL - Estado de cuenta con secciones (v1.1)
-- Cambio respecto a v1.0: se fijan en OPERATIVA MURCHISON las NC A-688, A-842 y A-870
-- (Nro interno 536, 860 y 931), que anulan facturas fijadas en esa hoja.
WITH G AS (
    SELECT
        X."ReconNum",
        MAX(CASE WHEN X."SrcObjTyp" = '24' THEN 1 ELSE 0 END)       AS "TieneCobro",
        MAX(CASE WHEN X."SrcObjTyp" = '14' THEN 1 ELSE 0 END)       AS "TieneNC",
        MAX(CASE
                WHEN X."ShortName" <> 'C00000137' THEN 1
                WHEN X."SrcObjTyp" NOT IN ('13', '14', '24', '203')
                     AND ABS(J."Debit" - J."Credit") > 0.005 THEN 1
                ELSE 0
            END)                                                    AS "TieneCanje",
        MAX(CASE WHEN X."SrcObjTyp" = '24' THEN J."RefDate" END)    AS "FechaCobro"
    FROM "ITR1" X
    INNER JOIN "OITR" R ON R."ReconNum" = X."ReconNum"
    INNER JOIN "JDT1" J ON J."TransId" = X."TransId"
                       AND J."Line_ID" = X."TransRowId"
    WHERE R."Canceled" = 'N'
    GROUP BY X."ReconNum"
),
P AS (
    SELECT
        I1."TransId",
        I1."TransRowId",
        SUM(I1."ReconSum")                                                AS "Aplicado",
        SUM(CASE WHEN G."TieneCobro" = 1 THEN I1."ReconSum" ELSE 0 END)   AS "Cobrado",
        SUM(CASE WHEN G."TieneCobro" = 0 AND G."TieneCanje" = 0 AND G."TieneNC" = 1
                 THEN I1."ReconSum" ELSE 0 END)                           AS "SoloNC",
        MAX(G."FechaCobro")                                               AS "FechaCobro"
    FROM "ITR1" I1
    INNER JOIN G ON G."ReconNum" = I1."ReconNum"
    WHERE I1."ShortName" = 'C00000137'
      AND I1."SrcObjTyp" <> '24'
    GROUP BY I1."TransId", I1."TransRowId"
),
D AS (
    SELECT
        I1."TransId",
        I1."TransRowId",
        STRING_AGG(
            CASE
                WHEN X."ShortName" <> I1."ShortName" THEN 'Canje con ' || X."ShortName"
                WHEN X."SrcObjTyp" = '24'  THEN 'Cobro'
                WHEN X."SrcObjTyp" = '14'  THEN 'NC'
                WHEN X."SrcObjTyp" = '13'  THEN 'Factura'
                WHEN X."SrcObjTyp" = '203' THEN 'Anticipo'
                ELSE 'Canje/asiento'
            END
            || ' ' || TO_VARCHAR(J."RefDate", 'DD/MM/YYYY')
            || ': ' || COALESCE(J."LineMemo", '')
            || ' (aplicado USD ' || TO_VARCHAR(TO_DECIMAL(I1."ReconSum", 18, 2)) || ')',
            ' | ' ORDER BY J."RefDate"
        ) AS "Detalle"
    FROM "ITR1" I1
    INNER JOIN G ON G."ReconNum" = I1."ReconNum"
    INNER JOIN "ITR1" X ON X."ReconNum" = I1."ReconNum"
                       AND NOT (X."TransId" = I1."TransId" AND X."TransRowId" = I1."TransRowId")
    INNER JOIN "JDT1" J ON J."TransId" = X."TransId"
                       AND J."Line_ID" = X."TransRowId"
    WHERE I1."ShortName" = 'C00000137'
      AND ABS(J."Debit" - J."Credit") > 0.005
      AND (   (G."TieneCobro" = 1 AND X."SrcObjTyp" = '24')
           OR (G."TieneCobro" = 0 AND (X."SrcObjTyp" <> I1."SrcObjTyp"
                                       OR X."ShortName" <> I1."ShortName")) )
    GROUP BY I1."TransId", I1."TransRowId"
),
L AS (
    SELECT '13' AS "ObjType", L1."DocEntry", H1."DocNum", LTRIM(L1."ItemCode", '0') AS "Art"
    FROM "INV1" L1
    INNER JOIN "OINV" H1 ON H1."DocEntry" = L1."DocEntry"
    WHERE H1."CardCode" = 'C00000137'
    UNION ALL
    SELECT '14', L2."DocEntry", H2."DocNum", LTRIM(L2."ItemCode", '0')
    FROM "RIN1" L2
    INNER JOIN "ORIN" H2 ON H2."DocEntry" = L2."DocEntry"
    WHERE H2."CardCode" = 'C00000137'
),
LS AS (
    SELECT
        "ObjType",
        "DocEntry",
        "DocNum",
        CASE
            WHEN "Art" IN ('11','12','13','14','15','16','36','62','68','70','71','77','89','96',
                           '137','159','317','318','348','349','350','351','352','366','402',
                           '463','546','567','568','784','942','1305','1369','1377','1411','1430')
                THEN 'ESCALAS'
            WHEN "Art" IN ('1246')
                THEN 'KAMENAR - CONFLUENCIA'
            WHEN "Art" IN ('155','1320','1321','1322','1323','1327','1343','1344','1398')
                THEN 'OPERATIVA MURCHISON'
            WHEN "Art" IN ('1247','1364','1439','1440')
                THEN 'REEMBARCOS'
            WHEN "Art" IN ('75','88','136','156','1056','1089','1119','1250','1289','1290',
                           '1317','1318','1319','1324','1326','1329','1330','1331','1332',
                           '1373','1378','1385',
                           '1090','1091','1092','1104',
                           '1291','1292','1293','1294','1295',
                           '1339')
                THEN 'OPERATIVA'
            ELSE 'SIN CLASIFICAR'
        END AS "Sec"
    FROM L
),
S AS (
    SELECT
        "ObjType",
        "DocEntry",
        CASE
            -- Historicos que ya estaban en OPERATIVA MURCHISON: quedan fijos (Nro interno de SAP)
            WHEN ("ObjType" = '13' AND "DocNum" IN (5251, 7911, 9839, 9895, 9896, 9898, 9899,
                                                    10112, 10404, 10405, 10406, 10407, 10409,
                                                    10410, 10685, 13677, 13678, 13679))
              OR ("ObjType" = '14' AND "DocNum" IN (572, 573, 574, 536, 860, 931))
                THEN 'OPERATIVA MURCHISON'
            WHEN SUM(CASE WHEN "Sec" = 'ESCALAS' THEN 1 ELSE 0 END) > 0 THEN 'ESCALAS'
            WHEN SUM(CASE WHEN "Sec" = 'KAMENAR - CONFLUENCIA' THEN 1 ELSE 0 END) > 0 THEN 'KAMENAR - CONFLUENCIA'
            WHEN SUM(CASE WHEN "Sec" = 'OPERATIVA MURCHISON' THEN 1 ELSE 0 END) > 0 THEN 'OPERATIVA MURCHISON'
            WHEN SUM(CASE WHEN "Sec" = 'REEMBARCOS' THEN 1 ELSE 0 END) > 0 THEN 'REEMBARCOS'
            WHEN SUM(CASE WHEN "Sec" = 'OPERATIVA' THEN 1 ELSE 0 END) > 0 THEN 'OPERATIVA'
            ELSE 'SIN CLASIFICAR'
        END AS "Seccion"
    FROM LS
    GROUP BY "ObjType", "DocEntry", "DocNum"
)
SELECT
    T0."RefDate"                                     AS "Fecha",
    CASE T0."TransType"
        WHEN 13  THEN 'Factura'
        WHEN 14  THEN 'Nota de credito'
        WHEN 203 THEN 'Factura anticipo'
        WHEN 24  THEN 'Cobro'
        WHEN 30  THEN 'Asiento'
        ELSE 'Otro'
    END                                              AS "Tipo",
    T0."LineMemo"                                    AS "Factura",
    T0."Debit" - T0."Credit"                         AS "Total",
    SUM(T0."Debit" - T0."Credit") OVER (
        ORDER BY T0."RefDate", T0."TransId", T0."Line_ID"
    )                                                AS "Saldo acumulado",
    T0."BalDueDeb" - T0."BalDueCred"                 AS "Saldo pendiente",
    COALESCE(T2."Project", T3."Project", T0."Project") AS "Proyecto",
    COALESCE(T2."Comments", T3."Comments", T4."Comments",
             T5."Comments", T1."Memo")               AS "Comentarios",
    CASE
        WHEN T0."TransType" IN (13, 14) THEN COALESCE(S."Seccion", 'SIN CLASIFICAR')
        ELSE ''
    END                                              AS "Seccion",
    CASE
        WHEN T0."TransType" NOT IN (13, 14, 203) THEN ''
        WHEN ABS(T0."BalDueDeb" - T0."BalDueCred") > 0.01 THEN
            CASE WHEN COALESCE(P."Aplicado", 0) > 0
                 THEN 'PENDIENTE PARCIAL' ELSE 'PENDIENTE' END
        WHEN COALESCE(P."SoloNC", 0) > 0
         AND COALESCE(P."Aplicado", 0) - COALESCE(P."SoloNC", 0) < 0.01 THEN 'ANULADA'
        ELSE 'PAGO'
    END                                              AS "Estado",
    COALESCE(P."Cobrado", 0)                         AS "Cobrado",
    P."FechaCobro"                                   AS "Fecha de cobro",
    CASE
        WHEN T0."TransType" NOT IN (13, 14, 203) THEN ''
        ELSE COALESCE(D."Detalle", 'Sin aplicar')
    END                                              AS "Detalle de pago"
FROM "JDT1" T0
INNER JOIN "OJDT" T1 ON T1."TransId" = T0."TransId"
LEFT JOIN "OINV" T2 ON T0."TransType" = 13  AND T2."DocEntry" = T0."CreatedBy"
LEFT JOIN "ORIN" T3 ON T0."TransType" = 14  AND T3."DocEntry" = T0."CreatedBy"
LEFT JOIN "ODPI" T4 ON T0."TransType" = 203 AND T4."DocEntry" = T0."CreatedBy"
LEFT JOIN "ORCT" T5 ON T0."TransType" = 24  AND T5."DocEntry" = T0."CreatedBy"
LEFT JOIN P ON P."TransId" = T0."TransId" AND P."TransRowId" = T0."Line_ID"
LEFT JOIN D ON D."TransId" = T0."TransId" AND D."TransRowId" = T0."Line_ID"
LEFT JOIN S ON S."ObjType" = T0."TransType" AND S."DocEntry" = T0."CreatedBy"
WHERE T0."ShortName" = 'C00000137'
ORDER BY T0."RefDate", T0."TransId", T0."Line_ID"
