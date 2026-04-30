"""
Domestic export item lookup using the annual export performance Excel file.
Queries 2025_연간 전체실적_전체.xlsx for Korean exports by country and product.
"""
import os
import re
import pandas as pd


# Countries that MAFRA considers under other ministry jurisdiction (→ write '-')
OTHER_MINISTRY_KEYWORDS = [
    '해수부', '식약처', '기후부', '환경부', '해양수산부', '식품의약품안전처',
    'fisheries', 'tobacco', '담배', 'water', '물',
]

# HS chapter → broad product category mapping
# Maps WTO product description keywords to HS chapter codes (2-digit)
PRODUCT_TO_HS_CHAPTERS = {
    # Animal products
    'poultry':       ['02'],
    'meat':          ['02'],
    'beef':          ['02'],
    'pork':          ['02'],
    'fish':          ['03', '16'],
    'seafood':       ['03', '16'],
    'dairy':         ['04'],
    'milk':          ['04'],
    'cheese':        ['04'],
    'egg':           ['04'],
    'hatching':      ['04'],
    'honey':         ['04'],
    'semen':         ['05'],
    'embryo':        ['05'],
    # Plant products
    'flower':        ['06'],
    'bulb':          ['06'],
    'plant':         ['06'],
    'cutting':       ['06'],
    'vegetable':     ['07'],
    'fruit':         ['08'],
    'apple':         ['08'],
    'pear':          ['08'],
    'grape':         ['08'],
    'citrus':        ['08'],
    'blueberry':     ['08'],
    'avocado':       ['08'],
    'cherry':        ['08'],
    'cereal':        ['10'],
    'wheat':         ['10'],
    'rice':          ['10'],
    'corn':          ['10'],
    'maize':         ['10'],
    'sorghum':       ['10'],
    'seed':          ['12'],
    'soybean':       ['12'],
    'cotton':        ['52'],
    'oil seed':      ['12'],
    'oilseed':       ['12'],
    'coffee':        ['09'],
    'tea':           ['09'],
    'spice':         ['09'],
    'mushroom':      ['07'],
    'ginseng':       ['12'],
    'wood':          ['44'],
    'timber':        ['44'],
    'log':           ['44'],
    'feed':          ['23'],
    'fodder':        ['23'],
    # Processed food
    'sugar':         ['17'],
    'chocolate':     ['18'],
    'cocoa':         ['18'],
    'flour':         ['11'],
    'bread':         ['19'],
    'pasta':         ['19'],
    'noodle':        ['19'],
    'sauce':         ['21'],
    'seasoning':     ['21'],
    'beverage':      ['22'],
    'alcohol':       ['22'],
    # Chemicals
    'pesticide':     ['38'],
    'fertilizer':    ['31'],
    'veterinary':    ['30'],
    'pharmaceutical':['30'],
}


class ExportLookup:
    def __init__(self):
        self._df = None
        self._path = None

    def load(self, path: str):
        """Load the export performance Excel into memory."""
        if self._df is not None and self._path == path:
            return  # Already loaded
        self._path = path
        # Load only needed columns for speed
        self._df = pd.read_excel(
            path,
            sheet_name='연간실적',
            usecols=['국가명', '구분', 'HSCODE', '품목명', '누계중량'],
            dtype={'HSCODE': str, '국가명': str, '품목명': str, '구분': str},
        )
        # Keep only export rows (구분 == 'E') with actual exports
        self._df = self._df[
            (self._df['구분'] == 'E') &
            (pd.to_numeric(self._df['누계중량'], errors='coerce').fillna(0) > 0)
        ].copy()
        self._df['HSCODE'] = self._df['HSCODE'].str.strip()
        self._df['국가명'] = self._df['국가명'].str.strip()
        self._df['품목명'] = self._df['품목명'].str.strip()

    def _hs_chapters_from_product(self, product_text: str) -> list:
        """Infer candidate HS chapters from product description keywords."""
        text_lower = product_text.lower()
        chapters = []
        for keyword, hs_list in PRODUCT_TO_HS_CHAPTERS.items():
            if keyword in text_lower:
                chapters.extend(hs_list)
        return list(set(chapters))

    def lookup(
        self,
        notifying_country: str,
        product_text: str,
        is_all_partners: bool,
        category: str = '',
    ) -> tuple:
        """
        Look up Korean export items for the given country and product.

        Trigger conditions (per manual):
          - is_all_partners=True: look up exports to all destinations,
            filter for the notifying country's exports
          - is_all_partners=False: look up exports to the specific notifying country

        Returns:
            (items_str, is_uncertain)
            items_str: comma-separated Korean product names, or '-'
            is_uncertain: True if the match required broad HS inference
        """
        if self._df is None:
            return ('-', False)

        # Under other ministry jurisdiction → '-'
        prod_lower = product_text.lower()
        if any(kw in prod_lower for kw in ['fish', 'seafood', 'tobacco', 'water supply']):
            return ('-', False)

        # Find matching country rows
        country_df = self._df[
            self._df['국가명'].str.lower() == notifying_country.strip().lower()
        ]
        if country_df.empty:
            # Try partial match
            country_df = self._df[
                self._df['국가명'].str.lower().str.contains(
                    notifying_country.strip().lower()[:4], na=False
                )
            ]

        if country_df.empty:
            return ('-', False)

        # Filter by HS chapter based on product description
        chapters = self._hs_chapters_from_product(product_text)
        is_uncertain = False

        if chapters:
            chapter_pattern = '|'.join(f'^{c}' for c in chapters)
            filtered = country_df[
                country_df['HSCODE'].str.match(chapter_pattern, na=False)
            ]
        else:
            # No clear HS hint → use broader category if available
            filtered = country_df
            is_uncertain = True

        if filtered.empty:
            return ('-', False)

        # Get unique product names, deduplicate
        items = filtered['품목명'].dropna().unique().tolist()
        items = list(dict.fromkeys(items))  # preserve order, deduplicate

        if not items:
            return ('-', False)

        # Limit to top 5 most relevant
        items = items[:5]
        return (', '.join(items), is_uncertain)

    def is_loaded(self):
        return self._df is not None


# Module-level singleton
_lookup = ExportLookup()


def get_lookup() -> ExportLookup:
    return _lookup
