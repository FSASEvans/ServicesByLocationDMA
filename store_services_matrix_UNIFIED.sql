-- ============================================================
-- STORE SERVICE MATRIX — UNIFIED (v4)
-- ============================================================
-- SINGLE SOURCE OF TRUTH: Pulls all transactions including
-- oil change visits. Oil change line items are classified as
-- L1 = 'Oil Change' so they are visible and separable.
-- Every row carries IS_OIL_TRANSACTION = 1 (oil change visit)
-- or 0 (standalone non-oil visit) as a filter flag.
--
-- The oil/non-oil split happens downstream (Excel tabs,
-- dashboard filter) — not at the query level.
--
-- OUTPUT FIELDS:
--   IS_OIL_TRANSACTION = 1 → line item came from a transaction
--                            that included an oil change
--   IS_OIL_TRANSACTION = 0 → pure non-oil visit
--   L1_CATEGORY = 'Oil Change' → the oil change line itself
--
-- v4 CHANGES (2026-03-26):
--   - BUGFIX: Removed TIE ROD from Engine L1 regex — it was
--     being caught before Suspension & Steering could evaluate
--     it. Tie Rods now correctly land in Suspension & Steering
--     L1 and Tie Rods L2.
--   - BUGFIX: Removed TIE ROD from Engine L2 regex for same
--     reason (Belt/Serpentine block also contained it).
--
-- REPLACE: store_services_matrix_UNIFIED_v3.sql (retired)
-- ============================================================

WITH

recent_transactions AS (
    SELECT TRANSACTION_ID, LOCATION_ID
    FROM CINCH_DATA.CINCH_TO_FULLSPEED_AUTOMOTIVE_DATA_SHARE.TRANSACTIONS
    WHERE TRANSACTION_AT >= DATEADD('month', -3, CURRENT_DATE())
),

-- Exclude only REC (Cinch recommendation) items — not oil changes
rec_only_txns AS (
    SELECT DISTINCT td.TRANSACTION_ID
    FROM CINCH_DATA.CINCH_TO_FULLSPEED_AUTOMOTIVE_DATA_SHARE.TRANSACTIONS_DETAILS AS td
    INNER JOIN recent_transactions AS rt ON td.TRANSACTION_ID = rt.TRANSACTION_ID
    WHERE td.ITEM_TYPE = 'REC'
),

-- Flag which transactions contain an oil change line item
-- (used to populate IS_OIL_TRANSACTION in output)
oil_txn_ids AS (
    SELECT DISTINCT td.TRANSACTION_ID
    FROM CINCH_DATA.CINCH_TO_FULLSPEED_AUTOMOTIVE_DATA_SHARE.TRANSACTIONS_DETAILS AS td
    INNER JOIN recent_transactions AS rt ON td.TRANSACTION_ID = rt.TRANSACTION_ID
    WHERE
        td.ITEM_TYPE IN ('OIL','OL1','OL6','OF','SYNTHETIC','OIL FILTERS',
                         'ENGINE OIL FS04','ENGINE OIL FS01')
        OR UPPER(td.ITEM_DESCR) LIKE '%OIL CHANGE%'
        OR UPPER(td.ITEM_DESCR) LIKE '%OIL & FILTER%'
        OR UPPER(td.ITEM_DESCR) LIKE '%OIL AND FILTER%'
        OR UPPER(td.ITEM_DESCR) LIKE '%OIL/FILTER%'
),

classified_lines AS (
    SELECT
        rt.LOCATION_ID,
        td.TRANSACTION_ID,
        td.ITEM                 AS SERVICE_CODE,
        td.ITEM_DESCR           AS SERVICE_NAME,

        -- ── L1 CATEGORY ─────────────────────────────────────────────────────
        CASE
            -- Oil Change — classified explicitly so it's visible and separable
            WHEN td.ITEM_TYPE IN ('OIL','OL1','OL6','OF','SYNTHETIC','OIL FILTERS',
                                  'ENGINE OIL FS04','ENGINE OIL FS01')
                 OR UPPER(td.ITEM_DESCR) LIKE '%OIL CHANGE%'
                 OR UPPER(td.ITEM_DESCR) LIKE '%OIL & FILTER%'
                 OR UPPER(td.ITEM_DESCR) LIKE '%OIL AND FILTER%'
                 OR UPPER(td.ITEM_DESCR) LIKE '%OIL/FILTER%'
                                                                THEN 'Oil Change'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(H1 |H3 |H4 |H7 |H7-|H8|H9|H10|H11|H13|9003|9004|9005|9006|9007|9008|9012|HALOGEN BULB|HEADLIGHT BULB).*' THEN 'Lighting'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(168 |194 |194N|906 |912 |921 |1156|1157|2057|2357|3057|3157|3457|3757|4157|6418|7440|7443|7444|LIGHT BULB|BULB).*' THEN 'Lighting'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%HEADLIGHT RESTORATION%' THEN 'Lighting'
            WHEN (UPPER(td.ITEM_DESCR) LIKE '%WIPER BLADE%' OR UPPER(td.ITEM_DESCR) LIKE '%BB PREMIUM WIPER%' OR UPPER(td.ITEM_DESCR) LIKE '%CB WIPER%' OR UPPER(td.ITEM_DESCR) RLIKE '.*\\d+-[12] REAR WIPER.*') THEN 'Wiper Blades'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(DHGOL|DHGO |DHSVR|DHGMP|AGM PLATM|AGM PLAT|GOLD DH|SILVER DH|H4 DIEHARD|BATTERY REPLACEMENT|DIEHARD GOLD).*' OR UPPER(td.ITEM_DESCR) RLIKE '^T[0-9]+ GOLD.*' THEN 'Battery'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%BATTERY%' THEN 'Battery'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(ALTERNATOR|STARTING & CHARGING|STARTER).*' THEN 'Battery'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(BRAKE|ROTOR|CALIPER).*' OR LOWER(td.ITEM_DESCR) IN ('front pads','rear pads','front brakes','rear brakes','pads','brake pads') OR UPPER(td.ITEM_DESCR) IN ('FRONT PADS','REAR PADS') THEN 'Brakes'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(TIRE REPLACEMENT|REPLACED.*TIRE|TIRE MOUNT|TIRE BALANCE|TIRE BALANCING|NEW TIRE|INSTALL.*TIRE).*' THEN 'Tire Sales'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(TIRE ROTATION|TIRE REPAIR|FLAT TIRE|TIRE DISPOSAL|TPMS|WHEEL BALANCE|TIRE INSPECTION).*' THEN 'Tire Services'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(BRAKE FLUID|DOT 3|DOT 4).*' THEN 'Fluids & Cooling'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(ANTIFREEZE|COOLANT|RADIATOR|POWER STEERING|WINDOW WASH|WINDSHIELD WASH).*' THEN 'Fluids & Cooling'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%COOLING SYSTEM LABOR%' THEN 'Fluids & Cooling'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(TRANSMISSION|TRANSFER CASE|CVT|DMX GLOBAL SYN ATF|DRIVETRAIN|HONDA DUAL PUMP FLUID|MERCON|DEXRON|AISIN ATF|\\bATF\\b|TRANS FLUID|TRANS SERV|TRANSMAX|PENNZOIL ATF).*' THEN 'Transmission'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(DIFFERENTIAL|GEARBOX|GEAR OIL|GEAR LUBE|AXLE OIL|AXLE FLUID|GL5|GL-5|75W-90|75W90|75W-140|75W140|80W-90|80W90|LIMITED SLIP|DIFF FLUID|DIFF OIL).*' THEN 'Differentials'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(OIL SYSTEM CLEAN|ENGINE CLEANER|ENGINE MAX|ENGINE TREATMENT|STOP LEAK|HIGH MILEAGE|\\bHMT\\b|FUEL SYSTEM CLEAN|FUEL JUELS|\\bFJ\\b|ENGINE FLUSH|\\bEFS\\b).*' THEN 'Additives'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(FUEL FILTER|GDI FUEL|INJECTOR CLEAN|THROTTLE BODY CLEAN|BG 44K|BG44K|FUEL INDUCTION|CARBON CLEAN).*' OR UPPER(td.ITEM_DESCR) RLIKE '^F[0-9]{5}.*' THEN 'Fuel System'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%CABIN AIR FILTER%' OR LOWER(td.ITEM_DESCR) IN ('cabin filter','cabin air filter') OR UPPER(td.ITEM_DESCR) RLIKE '^C[0-9]{5}.*' THEN 'Air Filters'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%AIR FILTER%' OR UPPER(td.ITEM_DESCR) RLIKE '^A[0-9]{4}.*' THEN 'Air Filters'
            -- ENGINE: TIE ROD removed — must reach Suspension & Steering below
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(SPARK PLUG|IGNITION COIL|IGNITION LABOR|COIL PAK|SERPENTINE|BELT TENSION|DRIVE BELT|VALVE COVER|THERMOSTAT|WATER PUMP|OIL PRESSURE|OIL PAN|MASS AIR FLOW|CV AXLE|ENGINE LABOR|ENGINE DIAGNOSTIC|COOLING SYSTEM LABOR|OVERLAP LABOR|EXHAUST LABOR|O2 SENSOR|02 SENSOR|OXYGEN SENSOR|DOWN STREAM|UP STREAM|BANK.*SENSOR|PCV VALVE|BLOWER MOTOR|PURGE VALVE|FUEL PUMP|INTAKE MANIFOLD|HEATER HOSE|MOTOR MOUNT|CLUTCH|WHEEL BEARING|BALL JOINT|CATALYTIC|HOSE REPLACE).*' THEN 'Engine'
            WHEN LOWER(td.ITEM_DESCR) IN ('belt','serp belt','idler pulley','sparkplugs','plugs','tune up','coil packs') THEN 'Engine'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(EMISSIONS|EMISSION CERTIFICATE|EMISSION STICKER|SAFETY STICKER|NC STATE|INSPECTION|VEHICLE CHECK|DIAGNOSTIC|SERVICE CHECKLIST|ON THE SPOT RENEWAL|UBER VEHICLE|FAILED NC STATE).*' THEN 'Emissions & Inspections'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(AIR CONDITION|REFRIGERANT|R-134|R134|A/C COMPRESSOR|AC COMPRESSOR).*' THEN 'Air Conditioning'
            -- SUSPENSION & STEERING: TIE ROD now correctly evaluated here
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(SUSPENSION|STEERING LABOR|LOWER CONTROL ARM|UPPER CONTROL ARM|SWAY BAR|STRUT|WHEEL HUB|TIE ROD|WHEEL STUD|ALIGNMENT).*' THEN 'Suspension & Steering'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(SHOP SUPPLIES|SHOP LABOR|MISC.*LABOR|LUG NUT|LABOR.*NO CHARGE|CAR WASH|SPECIAL FILTER|INSTALL).*' THEN 'Shop & Misc'
            WHEN LOWER(td.ITEM_DESCR) IN ('install','labor','cust own','top off') THEN 'Shop & Misc'
            ELSE 'Other'
        END AS L1_CATEGORY,

        -- ── L2 SUBCATEGORY ───────────────────────────────────────────────────
        CASE
            -- Oil Change L2
            WHEN td.ITEM_TYPE IN ('OIL','OL1','OL6','OF','SYNTHETIC','OIL FILTERS',
                                  'ENGINE OIL FS04','ENGINE OIL FS01')
                 OR UPPER(td.ITEM_DESCR) LIKE '%OIL CHANGE%'
                 OR UPPER(td.ITEM_DESCR) LIKE '%OIL & FILTER%'
                 OR UPPER(td.ITEM_DESCR) LIKE '%OIL AND FILTER%'
                 OR UPPER(td.ITEM_DESCR) LIKE '%OIL/FILTER%'              THEN 'Oil Change Service'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(H1 |H3 |H4 |H7 |H7-|H8|H9|H10|H11|H13|9003|9004|9005|9006|9007|9008|9012|HALOGEN BULB|HEADLIGHT BULB).*' THEN 'Headlight Bulbs'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(168 |194 |194N|906 |912 |921 |1156|1157|2057|2357|3057|3157|3457|3757|4157|6418|7440|7443|7444|LIGHT BULB|BULB).*' THEN 'Other Bulbs'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%HEADLIGHT RESTORATION%' THEN 'Headlight Restoration'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%WIPER BLADE%' OR UPPER(td.ITEM_DESCR) LIKE '%BB PREMIUM WIPER%' OR UPPER(td.ITEM_DESCR) LIKE '%CB WIPER%' OR UPPER(td.ITEM_DESCR) RLIKE '.*\\d+-[12] REAR WIPER.*' THEN 'Wiper Blades'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(DHGOL|DHGO |DHSVR|DHGMP|AGM PLATM|AGM PLAT|GOLD DH|SILVER DH|H4 DIEHARD|DIEHARD GOLD).*' OR UPPER(td.ITEM_DESCR) RLIKE '^T[0-9]+ GOLD.*' THEN 'Battery Replacement'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%BATTERY%' THEN 'Battery Replacement'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(ALTERNATOR|STARTING & CHARGING|STARTER).*' THEN 'Battery Testing'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(BRAKE|ROTOR|CALIPER).*' OR LOWER(td.ITEM_DESCR) IN ('front pads','rear pads','front brakes','rear brakes','pads','brake pads') THEN 'Brake Pads / Rotors'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(TIRE REPLACEMENT|REPLACED.*TIRE|TIRE MOUNT|TIRE BALANCE|TIRE BALANCING|NEW TIRE|INSTALL.*TIRE).*' THEN 'Tire Sales/Mount/Balance'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%TIRE DISPOSAL%' THEN 'Tire Disposal'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(TIRE ROTATION).*' THEN 'Tire Rotation'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(TIRE REPAIR|FLAT TIRE).*' THEN 'Tire Repair'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(TPMS).*' THEN 'TPMS'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(WHEEL BALANCE|TIRE INSPECTION).*' THEN 'Wheel Balance / Inspection'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(BRAKE FLUID|DOT 3|DOT 4).*' THEN 'Brake Fluid'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(ANTIFREEZE|COOLANT).*' THEN 'Coolant / Antifreeze Service'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(RADIATOR SERVICE|RADIATOR FLUSH|RADIATOR SERVICE).*' OR (UPPER(td.ITEM_DESCR) LIKE '%RADIATOR%' AND UPPER(td.ITEM_DESCR) LIKE '%SERVICE%') THEN 'Radiator Service'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(RADIATOR).*' AND UPPER(td.ITEM_DESCR) NOT LIKE '%SERVICE%' THEN 'Radiator Replacement'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%COOLING SYSTEM LABOR%' THEN 'Cooling System Labor'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%POWER STEERING%' THEN 'Power Steering Fluid'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(WINDOW WASH|WINDSHIELD WASH).*' THEN 'Window Wash'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(OIL SYSTEM CLEAN|\\bOSC\\b).*' THEN 'Oil System Cleaner'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(FUEL SYSTEM CLEAN|FUEL JUELS|\\bFJ\\b).*' THEN 'Fuel System Cleaner'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(ENGINE FLUSH|\\bEFS\\b).*' THEN 'Engine Cleaning Service'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(ENGINE CLEANER|ENGINE MAX|ENGINE TREATMENT|\\bEMP\\b|\\bEMQ\\b).*' THEN 'Engine Treatment'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(STOP LEAK|\\bESL\\b).*' THEN 'Stop Leak'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(HIGH MILEAGE|\\bHMT\\b).*' THEN 'High Mileage Treatment'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(TRANSMISSION FLUID|TRANS FLUID|TRANS SERV|CVT|DMX GLOBAL SYN ATF|HONDA DUAL PUMP|MERCON|DEXRON|AISIN ATF|TRANSMAX|PENNZOIL ATF|\\bATF\\b).*' THEN 'Transmission Fluid Exchange'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%TRANSFER CASE%' THEN 'Transfer Case'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(DRIVETRAIN LABOR|TRANSMISSION LABOR).*' THEN 'Drivetrain Labor'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%FRONT DIFFERENTIAL%' THEN 'Front Differential'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%REAR DIFFERENTIAL%' THEN 'Rear Differential'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(GL5|GL-5|GEAR OIL|GEAR LUBE|AXLE OIL|AXLE FLUID|75W|80W|LIMITED SLIP|DIFF FLUID|DIFF OIL|GEARBOX).*' THEN 'Gear Oil'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(GDI FUEL|INJECTOR CLEAN|THROTTLE BODY CLEAN|BG 44K|BG44K|FUEL INDUCTION|CARBON CLEAN).*' THEN 'Fuel System Cleaning'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%FUEL FILTER%' OR UPPER(td.ITEM_DESCR) RLIKE '^F[0-9]{5}.*' THEN 'Fuel Filter'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%CABIN AIR FILTER%' OR LOWER(td.ITEM_DESCR) IN ('cabin filter','cabin air filter') OR UPPER(td.ITEM_DESCR) RLIKE '^C[0-9]{5}.*' THEN 'Cabin Air Filter'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%AIR FILTER%' OR UPPER(td.ITEM_DESCR) RLIKE '^A[0-9]{4}.*' THEN 'Engine Air Filter'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(SPARK PLUG|COIL PAK).*' OR LOWER(td.ITEM_DESCR) IN ('sparkplugs','plugs','tune up') THEN 'Spark Plugs'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(IGNITION COIL|IGNITION LABOR|COIL PACK).*' THEN 'Ignition / Coils'
            -- BELTS: TIE ROD removed from this block
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(SERPENTINE|BELT TENSION|DRIVE BELT|COOLING SYSTEM LABOR|OVERLAP LABOR|EXHAUST LABOR).*' OR LOWER(td.ITEM_DESCR) IN ('belt','serp belt','idler pulley') THEN 'Belts'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(ENGINE LABOR|ENGINE DIAGNOSTIC).*' THEN 'Engine Labor'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(CLUTCH|CYLINDER REPLACE).*' THEN 'Clutch / Cylinder'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(WHEEL BEARING|WHEEL HUB BEARING).*' THEN 'Wheel Bearings'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(BALL JOINT).*' THEN 'Ball Joints'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(CATALYTIC|CATALYTIC CONVERTER).*' THEN 'Catalytic Converter'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(HOSE REPLACE|RADIATOR HOSE|HEATER HOSE).*' THEN 'Hose Replacement'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%VALVE COVER%' THEN 'Valve Cover Gasket'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(THERMOSTAT|WATER PUMP).*' THEN 'Thermostat / Water Pump'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(O2 SENSOR|02 SENSOR|OXYGEN SENSOR|DOWN STREAM|UP STREAM|BANK.*SENSOR|PCV VALVE).*' THEN 'O2 / Sensors'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(EMISSIONS|EMISSION CERTIFICATE|EMISSION STICKER|SAFETY STICKER|NC STATE|FAILED NC STATE|ON THE SPOT RENEWAL).*' THEN 'Emissions Test'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(VEHICLE CHECK|DIAGNOSTIC|SERVICE CHECKLIST|INSPECTION|UBER VEHICLE).*' THEN 'Vehicle Check / Inspection'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(AIR CONDITION|REFRIGERANT|R-134|R134|A/C COMPRESSOR|AC COMPRESSOR).*' THEN 'A/C Service'
            -- TIE RODS: now reachable — no longer shadowed by Engine block above
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(TIE ROD).*' THEN 'Tie Rods'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(SUSPENSION|STEERING LABOR|LOWER CONTROL ARM|UPPER CONTROL ARM|SWAY BAR|STRUT|WHEEL HUB|WHEEL STUD).*' THEN 'Suspension / Steering Labor'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%SHOP SUPPLIES%' THEN 'Shop Supplies'
            WHEN UPPER(td.ITEM_DESCR) RLIKE '.*(SHOP LABOR|MISC.*LABOR|LABOR.*NO CHARGE).*' THEN 'Shop Labor'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%LUG NUT%' THEN 'Lug Nut Service'
            WHEN UPPER(td.ITEM_DESCR) LIKE '%CAR WASH%' THEN 'Car Wash'
            ELSE 'Uncategorized'
        END AS L2_SUBCATEGORY

    FROM recent_transactions AS rt
    INNER JOIN CINCH_DATA.CINCH_TO_FULLSPEED_AUTOMOTIVE_DATA_SHARE.TRANSACTIONS_DETAILS AS td
        ON rt.TRANSACTION_ID = td.TRANSACTION_ID
    WHERE
        -- Only exclude REC items — oil change transactions are KEPT
        td.TRANSACTION_ID NOT IN (SELECT TRANSACTION_ID FROM rec_only_txns)
        AND td.ITEM_DESCR IS NOT NULL
        AND TRIM(td.ITEM_DESCR) <> ''
        AND COALESCE(td.ITEM_TYPE, '') NOT IN ('n/a', '')
        AND td.ITEM_TYPE IS NOT NULL
),

with_store AS (
    SELECT
        l.ENTITY_NUMBER             AS STORE_NUMBER,
        l.FRIENDLY_NAME             AS STORE_NAME,
        l.NAME                      AS STORE_NAME_FULL,
        l.CITY,
        l.STATE,
        l.LATITUDE,
        l.LONGITUDE,
        l.ADDRESS1,
        l.ZIP_CODE,
        cl.TRANSACTION_ID,
        cl.SERVICE_CODE,
        cl.SERVICE_NAME,
        cl.L1_CATEGORY,
        cl.L2_SUBCATEGORY,
        -- Flag whether this transaction was an oil change visit
        CASE WHEN ot.TRANSACTION_ID IS NOT NULL THEN 1 ELSE 0 END AS IS_OIL_TRANSACTION
    FROM classified_lines AS cl
    INNER JOIN CINCH_DATA.CINCH_TO_FULLSPEED_AUTOMOTIVE_DATA_SHARE.LOCATIONS AS l
        ON cl.LOCATION_ID = l.LOCATION_ID
    LEFT JOIN oil_txn_ids AS ot
        ON cl.TRANSACTION_ID = ot.TRANSACTION_ID
    WHERE l.IS_ACTIVE = TRUE
),

final AS (
    SELECT
        STORE_NUMBER,
        STORE_NAME,
        STORE_NAME_FULL,
        CITY,
        STATE,
        LATITUDE,
        LONGITUDE,
        ADDRESS1,
        ZIP_CODE,
        L1_CATEGORY,
        L2_SUBCATEGORY,
        SERVICE_CODE,
        SERVICE_NAME,
        IS_OIL_TRANSACTION,
        COUNT(DISTINCT TRANSACTION_ID)  AS TRANSACTION_COUNT,
        1                               AS OFFERS_SERVICE
    FROM with_store
    WHERE STORE_NUMBER IS NOT NULL
    GROUP BY
        STORE_NUMBER, STORE_NAME, STORE_NAME_FULL,
        CITY, STATE, LATITUDE, LONGITUDE, ADDRESS1, ZIP_CODE,
        L1_CATEGORY, L2_SUBCATEGORY, SERVICE_CODE, SERVICE_NAME,
        IS_OIL_TRANSACTION
)

SELECT * FROM final
ORDER BY STORE_NUMBER, IS_OIL_TRANSACTION DESC, L1_CATEGORY, L2_SUBCATEGORY, SERVICE_NAME;
