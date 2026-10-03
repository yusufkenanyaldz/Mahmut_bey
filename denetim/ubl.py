"""UBL-TR (e-Fatura / e-Arşiv) XML ayrıştırıcı."""
import xml.etree.ElementTree as ET

from .utils import normalize_vkn, parse_date

NS = {
    "inv": "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2",
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
}
VAT_CODE = "0015"  # KDV vergi türü kodu

# GİB vergi türü kodları (KDV dışındaki vergilerin raporda okunur adı için)
VERGI_ADLARI = {
    "0003": "GV Stopajı", "0011": "KV Stopajı", "0021": "BSMV", "0059": "Konaklama Vergisi",
    "0061": "KKDF Kesintisi", "0071": "ÖTV I (Petrol)", "0073": "ÖTV III (Kolalı Gazoz)",
    "0074": "ÖTV III (Alkollü İçecek)", "0075": "ÖTV III (Tütün)", "0076": "ÖTV III (Puro/Sigara)",
    "0077": "ÖTV IV (Dayanıklı Tüketim)", "1047": "Damga Vergisi", "1048": "5035 Damga Vergisi",
    "4071": "Elektrik Havagazı Tüketim Vergisi", "4080": "ÖİV", "4081": "5035 ÖİV", "8001": "Borsa Tescil Ücreti",
    "8002": "Enerji Fonu", "8004": "TRT Payı", "8005": "Elektrik Tüketim Vergisi", "8006": "Telsiz Kullanım Ücreti",
    "8007": "Telsiz Ruhsat Ücreti", "8008": "Çevre Temizlik Vergisi", "9021": "4961 BSMV",
    "9077": "ÖTV II (Motorlu Taşıtlar)", "9944": "Hal Rüsumu",
}


class UBLParseError(Exception):
    pass


def _text(node, path):
    if node is None:
        return None
    el = node.find(path, NS)
    if el is None or el.text is None:
        return None
    return el.text.strip() or None


def _num(node, path, default=None):
    value = _text(node, path)
    if value is None:
        return default
    try:
        return float(value)
    except ValueError:
        raise UBLParseError(f"{path} alanı sayı değil: {value}") from None


def _party(party_node):
    """(vkn/tckn, unvan) döndürür."""
    if party_node is None:
        return "", ""
    party = party_node.find("cac:Party", NS)
    if party is None:
        return "", ""
    vkn = ""
    for pid in party.findall("cac:PartyIdentification/cbc:ID", NS):
        if (pid.get("schemeID") or "").upper() in ("VKN", "TCKN") and pid.text:
            vkn = normalize_vkn(pid.text)
            break
    name = _text(party, "cac:PartyName/cbc:Name")
    if not name:
        first = _text(party, "cac:Person/cbc:FirstName") or ""
        family = _text(party, "cac:Person/cbc:FamilyName") or ""
        name = f"{first} {family}".strip()
    return vkn, name or ""


def _diger_vergiler(subtotals):
    """KDV (0015) dışındaki vergi alt toplamlarını vergi türü koduna göre toplar.

    Dönüş: [{"code", "name", "amount", "percent", "taxable"}] (belge para biriminde, koda göre sıralı)
    """
    out = {}
    for sub in subtotals:
        code = _text(sub, "cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode")
        if not code or code == VAT_CODE:
            continue
        t = out.setdefault(code, {"code": code, "name": VERGI_ADLARI.get(code)
                                  or _text(sub, "cac:TaxCategory/cac:TaxScheme/cbc:Name") or code,
                                  "amount": 0.0, "percent": _num(sub, "cbc:Percent", None), "taxable": 0.0})
        t["amount"] += _num(sub, "cbc:TaxAmount", 0.0)
        t["taxable"] += _num(sub, "cbc:TaxableAmount", 0.0)
    for t in out.values():
        t["amount"], t["taxable"] = round(t["amount"], 2), round(t["taxable"], 2)
    return [out[k] for k in sorted(out)]


def _tevkifat(totals):
    """WithholdingTaxTotal (KDV tevkifatı) tutarı, oranı (%) ve tevkifat kodu."""
    amount, rate, code = 0.0, None, None
    for tot in totals:
        subs = tot.findall("cac:TaxSubtotal", NS)
        if subs:
            for sub in subs:
                amount += _num(sub, "cbc:TaxAmount", 0.0)
                rate = rate if rate is not None else _num(sub, "cbc:Percent", None)
                code = code or _text(sub, "cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode")
        else:
            amount += _num(tot, "cbc:TaxAmount", 0.0)
    return {"withholding_amount": round(amount, 2), "withholding_rate": rate, "withholding_code": code}


def _find_invoice_root(root):
    if root.tag == f"{{{NS['inv']}}}Invoice":
        return root
    # Bazı entegratörler faturayı bir zarf (envelope) içinde gönderir
    found = root.find(f".//{{{NS['inv']}}}Invoice")
    if found is None:
        raise UBLParseError("Dosyada UBL-TR Invoice öğesi bulunamadı")
    return found


def parse_ubl(content, source_file=None):
    """XML içeriğini (bytes) ayrıştırır.

    Dönüş: (header, lines, warnings)
    """
    try:
        root = ET.fromstring(content)
    except ET.ParseError as e:
        raise UBLParseError(f"XML okunamadı: {e}") from None
    inv = _find_invoice_root(root)
    warnings = []

    invoice_no = _text(inv, "cbc:ID")
    if not invoice_no:
        raise UBLParseError("Fatura numarası (cbc:ID) yok")
    try:
        issue_date = parse_date(_text(inv, "cbc:IssueDate"))
    except ValueError as e:
        raise UBLParseError(f"Fatura tarihi hatalı: {e}") from None

    supplier_vkn, supplier_name = _party(inv.find("cac:AccountingSupplierParty", NS))
    customer_vkn, customer_name = _party(inv.find("cac:AccountingCustomerParty", NS))
    if not supplier_vkn:
        raise UBLParseError("Satıcı VKN/TCKN bulunamadı")

    currency = (_text(inv, "cbc:DocumentCurrencyCode") or "TRY").upper()
    if currency == "TL":
        currency = "TRY"
    rate = 1.0
    if currency != "TRY":
        rate = _num(inv, "cac:PricingExchangeRate/cbc:CalculationRate", None)
        if not rate:
            warnings.append(f"{currency} faturada döviz kuru (PricingExchangeRate) yok, kur 1 kabul edildi")
            rate = 1.0

    monetary = inv.find("cac:LegalMonetaryTotal", NS)
    if monetary is None:
        raise UBLParseError("LegalMonetaryTotal (parasal toplamlar) bölümü yok")
    line_ext = _num(monetary, "cbc:LineExtensionAmount", 0.0)
    allowance = _num(monetary, "cbc:AllowanceTotalAmount", 0.0)
    charge = _num(monetary, "cbc:ChargeTotalAmount", 0.0)
    tax_excl = _num(monetary, "cbc:TaxExclusiveAmount", None)
    if tax_excl is None:
        tax_excl = line_ext - allowance + charge
    payable = _num(monetary, "cbc:PayableAmount", None)

    vat_total = 0.0
    for sub in inv.findall("cac:TaxTotal/cac:TaxSubtotal", NS):
        if _text(sub, "cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode") == VAT_CODE:
            vat_total += _num(sub, "cbc:TaxAmount", 0.0)
    # KDV dışındaki vergiler (ÖTV, ÖİV, konaklama ...): belge toplamında yoksa satırlardan toplanır
    taxes = _diger_vergiler(inv.findall("cac:TaxTotal/cac:TaxSubtotal", NS))
    if not taxes:
        taxes = _diger_vergiler(inv.findall("cac:InvoiceLine/cac:TaxTotal/cac:TaxSubtotal", NS))
    # KDV tevkifatı: belge toplamında yoksa satırlardaki WithholdingTaxTotal toplanır
    withholding = _tevkifat(inv.findall("cac:WithholdingTaxTotal", NS))
    if withholding["withholding_amount"] == 0:
        withholding = _tevkifat(inv.findall("cac:InvoiceLine/cac:WithholdingTaxTotal", NS))

    header = {
        "invoice_no": invoice_no,
        "issue_date": issue_date,
        "supplier_vkn": supplier_vkn,
        "supplier_name": supplier_name,
        "customer_vkn": customer_vkn,
        "customer_name": customer_name,
        "invoice_type": (_text(inv, "cbc:InvoiceTypeCode") or "SATIS").upper(),
        "profile": _text(inv, "cbc:ProfileID"),
        "currency": currency,
        "exchange_rate": rate,
        "line_extension_amount": line_ext,
        "allowance_total": allowance,
        "total_amount": tax_excl,
        "vat_amount": vat_total,
        "payable_amount": payable,
        "taxes": taxes,
        "other_tax_amount": round(sum(t["amount"] for t in taxes), 2),
        **withholding,
        "tax_detail": 1,
        "source": "XML",
        "source_file": source_file,
    }

    lines = []
    for idx, ln in enumerate(inv.findall("cac:InvoiceLine", NS), start=1):
        qty_el = ln.find("cbc:InvoicedQuantity", NS)
        try:
            qty = float(qty_el.text) if qty_el is not None and qty_el.text else 0.0
        except ValueError:
            raise UBLParseError(f"{idx}. satır miktarı sayı değil") from None
        line_net = _num(ln, "cbc:LineExtensionAmount", 0.0)
        name = _text(ln, "cac:Item/cbc:Name") or _text(ln, "cac:Item/cbc:Description") or "(adsız ürün)"
        vat_rate, vat_amount = None, 0.0
        for sub in ln.findall("cac:TaxTotal/cac:TaxSubtotal", NS):
            if _text(sub, "cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode") == VAT_CODE:
                vat_rate = _num(sub, "cbc:Percent", None)
                vat_amount += _num(sub, "cbc:TaxAmount", 0.0)
        if qty == 0:
            warnings.append(f"{idx}. satırda ({name}) miktar 0, fiyat analizine alınmayacak")
        lines.append({
            "line_no": int(_text(ln, "cbc:ID") or idx) if (_text(ln, "cbc:ID") or "").isdigit() else idx,
            "item_name": name,
            "quantity": qty,
            "uom": (qty_el.get("unitCode") if qty_el is not None else None) or "ADET",
            "unit_price": _num(ln, "cac:Price/cbc:PriceAmount", None),
            "line_net": line_net,
            "vat_rate": vat_rate,
            "vat_amount": vat_amount,
        })
    if not lines:
        warnings.append("Faturada hiç satır (InvoiceLine) yok")
    return header, lines, warnings
