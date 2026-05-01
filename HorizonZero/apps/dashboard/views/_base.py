import json

MESES_ES = {
    1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr',
    5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Ago',
    9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic',
}


def build_timeline_heatmap(raw_timeline: list) -> tuple:
    timeline_data = [
        {"label": f"{MESES_ES[r['mes']]} {r['anio']}", "total": r["total"]}
        for r in raw_timeline
    ]
    heatmap_dict  = {}
    heatmap_anios = []
    for r in raw_timeline:
        a, m, t = r["anio"], r["mes"], r["total"]
        if a not in heatmap_dict:
            heatmap_dict[a] = {}
            heatmap_anios.append(a)
        heatmap_dict[a][m] = t
    return timeline_data, heatmap_anios, json.dumps(heatmap_dict)
