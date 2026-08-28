@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def consultar_cache(url: str, params_tuple: Tuple[Tuple[str, str], ...], max_paginas: int) -> Tuple[List[Dict[str, Any]], int, Optional[int]]:
    base = dict(params_tuple)
    base["pagina"] = 1
    
    primeira = get_json(url, base)
    regs1 = registros_api(primeira)
    if not regs1:
        return [], 0, paginacao(primeira).get("totalPaginas")
        
    info = paginacao(primeira)
    total = info.get("totalPaginas")
    
    if total is None and info.get("totalRegistros") is not None:
        tamanho = max(1, int(base.get("tamanhoPagina", 50)))
        total = (info["totalRegistros"] + tamanho - 1) // tamanho
        
    limite = max(1, min(max_paginas, total or max_paginas))
    todos = list(regs1)
    
    if limite == 1:
        return todos, 1, total

    # CORREÇÃO: Usa a função get_json padrão que já utiliza sessao_http() global
    def buscar(pagina: int):
        p = dict(base)
        p["pagina"] = pagina
        data = get_json(url, p)
        return pagina, registros_api(data)

    resultados = {}
    # Usa a ThreadPool para buscar as demais páginas simultaneamente
    with ThreadPoolExecutor(max_workers=min(MAX_WORKERS, limite - 1)) as pool:
        futuros = {pool.submit(buscar, pagina): pagina for pagina in range(2, limite + 1)}
        for f in as_completed(futuros):
            pagina = futuros[f]
            try:
                resultados[pagina] = f.result()[1]
            except Exception:
                resultados[pagina] = []
                
    for pagina in range(2, limite + 1):
        todos.extend(resultados.get(pagina, []))
        
    return todos, limite, total
