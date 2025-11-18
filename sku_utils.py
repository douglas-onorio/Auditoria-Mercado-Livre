import pandas as pd
import re

def aplicar_custos(df, df_custos, coluna_unidades):
    """
    Regras implementadas:
    ✔ Se SKU existir exatamente no df_custos → usa esse SKU direto.
    ✔ Se SKU terminar em C2..C12 → custo composto real → usa ele direto.
    ✔ Se SKU tiver hífen e NÃO existir no df_custos → explode (é pacote do ML).
    ✔ Pacote 'mãe' (Agrupado - Pacotes) nunca recebe custo.
    ✔ Itens filhos usam custo normal por SKU simples.
    """

    # === NORMALIZA DF_CUSTOS (mantém hífen sempre) ===
    df_custos = df_custos.copy()
    df_custos["SKU"] = df_custos["SKU"].astype(str).str.strip()

    # Dicionário rápido de lookup
    custos_map = dict(zip(df_custos["SKU"], df_custos["Custo_Produto"]))

    def detectar_composto_cX(sku):
        """
        Ex: 3990C2 → (base=3990, multiplicador=2)
        Só válido se final for C2..C12
        """
        sku = sku.upper().strip()
        match = re.match(r"^(\d+)[Cc]([2-9]|1[0-2])$", sku)
        if match:
            base = match.group(1)
            mult = int(match.group(2))
            return base, mult
        return None

    custos_totais = []
    custos_unitarios = []

    for idx, row in df.iterrows():
        sku = str(row.get("SKU", "")).strip()
        tipo = str(row.get("Tipo_Anuncio", "")).lower()
        unidades = row.get(coluna_unidades, 1)

        # === CASO 1 — Pacote mãe: custo sempre ZERO ===
        if "agrupado (pacotes" in tipo:
            custos_unitarios.append(0.0)
            custos_totais.append(0.0)
            continue

        # === CASO 2 — SKU está na planilha → usa esse SKU direto ===
        if sku in custos_map:
            custo_unit = float(custos_map[sku])
            custos_unitarios.append(custo_unit)
            custos_totais.append(round(custo_unit * unidades, 2))
            continue

        # === CASO 3 — SKU COMPOSTO REAL (C2..C12) ===
        composto = detectar_composto_cX(sku)
        if composto:
            base, mult = composto
            if base in custos_map:
                custo_base = float(custos_map[base])
                custo_unit = round(custo_base * mult, 2)
                custos_unitarios.append(custo_unit)
                custos_totais.append(round(custo_unit * unidades, 2))
                continue

        # === CASO 4 — SKU com hífen → pacote filho OU composição do ML ===
        if "-" in sku:
            partes = [p.strip() for p in sku.split("-") if p.strip()]

            # explode cada parte
            custo_total = 0
            custo_unit_sku = 0

            for p in partes:
                # composto dentro de hífen (caso raro)
                comp = detectar_composto_cX(p)
                if comp:
                    base, mult = comp
                    if base in custos_map:
                        custo_total += float(custos_map[base]) * mult
                        continue

                if p in custos_map:
                    custo_total += float(custos_map[p])

            custo_total = round(custo_total * unidades, 2)
            custo_unit_sku = round(custo_total / unidades, 2) if unidades else custo_total

            custos_unitarios.append(custo_unit_sku)
            custos_totais.append(custo_total)
            continue

        # === CASO 5 — SKU simples normal ===
        if sku in custos_map:
            custo_unit = float(custos_map[sku])
            custos_unitarios.append(custo_unit)
            custos_totais.append(round(custo_unit * unidades, 2))
        else:
            # SKU totalmente desconhecido
            custos_unitarios.append(0.0)
            custos_totais.append(0.0)

    df["Custo_Produto_Unitario"] = custos_unitarios
    df["Custo_Produto_Total"] = custos_totais
    return df
