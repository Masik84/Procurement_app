DO $$
BEGIN
    IF EXISTS (
        SELECT 1
        FROM information_schema.columns
        WHERE table_schema = 'public'
          AND table_name = 'target_price_calculations'
          AND column_name = 'cost_novo_wvat_recalculated'
    ) AND NOT EXISTS (
        SELECT 1
        FROM information_schema.columns
        WHERE table_schema = 'public'
          AND table_name = 'target_price_calculations'
          AND column_name = 'cost_novo_wvat'
    ) THEN
        ALTER TABLE public.target_price_calculations
        RENAME COLUMN cost_novo_wvat_recalculated TO cost_novo_wvat;
    END IF;
END $$;

ALTER TABLE public.target_price_calculations
ADD COLUMN IF NOT EXISTS donor_supplier_price NUMERIC NULL;

ALTER TABLE public.target_price_calculations
ADD COLUMN IF NOT EXISTS donor_currency_code VARCHAR(10) NULL;

ALTER TABLE public.target_price_calculations
ADD COLUMN IF NOT EXISTS donor_fx_rate_used NUMERIC NULL;
