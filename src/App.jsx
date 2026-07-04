import { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

// ─── Dados ────────────────────────────────────────────────────────────────────
const ITEMS = [
  {id:0, cat:'BOMBONIERE', nome:'BALA HALL\'S (preto, vermelho e azul)'},
  {id:1, cat:'BOMBONIERE', nome:'CHICLE TRIDENT (verde, rosa e azul)'},
  {id:2, cat:'BOMBONIERE', nome:'CHOC BARRA KILO meio amargo'},
  {id:3, cat:'BOMBONIERE', nome:'CHOC LAKA'},
  {id:4, cat:'BOMBONIERE', nome:'CHOC PRESTIGIO'},
  {id:5, cat:'BOMBONIERE', nome:'CHOC SONHO DE VALSA'},
  {id:6, cat:'BOMBONIERE', nome:'CHOC TALENTO (verde, vermelho e roxo)'},
  {id:7, cat:'BEBIDAS COM ÁLCOOL', nome:'AGUARD. SALINAS'},
  {id:8, cat:'BEBIDAS COM ÁLCOOL', nome:'CAMPARI'},
  {id:9, cat:'BEBIDAS COM ÁLCOOL', nome:'COINTREAU'},
  {id:10, cat:'BEBIDAS COM ÁLCOOL', nome:'CONHAQUE DOMECQ'},
  {id:11, cat:'BEBIDAS COM ÁLCOOL', nome:'CURAÇAU BLUE'},
  {id:12, cat:'BEBIDAS COM ÁLCOOL', nome:'LICOR AMARULA'},
  {id:13, cat:'BEBIDAS COM ÁLCOOL', nome:'MARTINI'},
  {id:14, cat:'BEBIDAS COM ÁLCOOL', nome:'RUM BACARDI'},
  {id:15, cat:'BEBIDAS COM ÁLCOOL', nome:'TEQUILA JOSÉ CUERVO'},
  {id:16, cat:'BEBIDAS COM ÁLCOOL', nome:'VINHO CANÇÃO TTO SVE'},
  {id:17, cat:'BEBIDAS COM ÁLCOOL', nome:'VINHO SANTA FELICIDADE TTO SVE'},
  {id:18, cat:'BEBIDAS COM ÁLCOOL', nome:'VODKA ORLOFF'},
  {id:19, cat:'BEBIDAS COM ÁLCOOL', nome:'WHISKY JOHNNIE WALKER RED'},
  {id:20, cat:'BEBIDAS COM ÁLCOOL', nome:'WHISKY JOHNNIE WALKER BLACK'},
  {id:21, cat:'CERVEJAS', nome:'BRAHMA 600 ML'},
  {id:22, cat:'CERVEJAS', nome:'BRAHMA LATA 350 ML'},
  {id:23, cat:'CERVEJAS', nome:'HEINEKEN 600 ML'},
  {id:24, cat:'CERVEJAS', nome:'HEINEKEN LATA 350 ML'},
  {id:25, cat:'CERVEJAS', nome:'Heineken long neck'},
  {id:26, cat:'CERVEJAS', nome:'Heineken Zero lata'},
  {id:27, cat:'CERVEJAS', nome:'ORIGINAL 600 ML'},
  {id:28, cat:'BEBIDAS SEM ÁLCOOL', nome:'ÁGUA MINERAL BONISSIMA COM GÁS 500 ML'},
  {id:29, cat:'BEBIDAS SEM ÁLCOOL', nome:'ÁGUA MINERAL BONISSIMA SEM GÁS 500 ML'},
  {id:30, cat:'BEBIDAS SEM ÁLCOOL', nome:'Água Tônica 350 ml'},
  {id:31, cat:'BEBIDAS SEM ÁLCOOL', nome:'Coca 2L'},
  {id:32, cat:'BEBIDAS SEM ÁLCOOL', nome:'Coca lata 350ML'},
  {id:33, cat:'BEBIDAS SEM ÁLCOOL', nome:'Coca Zero 2L'},
  {id:34, cat:'BEBIDAS SEM ÁLCOOL', nome:'Coca Zero lata 350ML'},
  {id:35, cat:'BEBIDAS SEM ÁLCOOL', nome:'Fanta lata 350ML'},
  {id:36, cat:'BEBIDAS SEM ÁLCOOL', nome:'H2OH 500ml'},
  {id:37, cat:'BEBIDAS SEM ÁLCOOL', nome:'Mate couro 1 L'},
  {id:38, cat:'BEBIDAS SEM ÁLCOOL', nome:'Mate Couro 2 L'},
  {id:39, cat:'BEBIDAS SEM ÁLCOOL', nome:'Sprite lata 350ML'},
  {id:40, cat:'BEBIDAS SEM ÁLCOOL', nome:'Suco lata DELVALLE Maracujá'},
  {id:41, cat:'BEBIDAS SEM ÁLCOOL', nome:'Suco lata DELVALLE Pêssego'},
  {id:42, cat:'BEBIDAS SEM ÁLCOOL', nome:'Suco lata DELVALLE Uva'},
  {id:43, cat:'BEBIDAS SEM ÁLCOOL', nome:'Suco Caju Maguary'},
  {id:44, cat:'BEBIDAS SEM ÁLCOOL', nome:'Suco Maracujá Maguary'},
  {id:45, cat:'BEBIDAS SEM ÁLCOOL', nome:'Suco Tangerina em pó Vilma 1 kg'},
  {id:46, cat:'MERCEARIA', nome:'Açúcar Cristal 5KG'},
  {id:47, cat:'MERCEARIA', nome:'Açúcar refinado 1 kg'},
  {id:48, cat:'MERCEARIA', nome:'Açúcar SACHET'},
  {id:49, cat:'MERCEARIA', nome:'AJI-NO-MOTO 25 KG'},
  {id:50, cat:'MERCEARIA', nome:'Amendoim torrado moído pacha 500g'},
  {id:51, cat:'MERCEARIA', nome:'Arroz CAMIL ou Tio João 5 KG'},
  {id:52, cat:'MERCEARIA', nome:'ATUM COQUEIRO SÓLIDO'},
  {id:53, cat:'MERCEARIA', nome:'AZEITE 200 ML LATA'},
  {id:54, cat:'MERCEARIA', nome:'BANHA SADIA 30 KG'},
  {id:55, cat:'MERCEARIA', nome:'Café 3 corações 250g'},
  {id:56, cat:'MERCEARIA', nome:'CALDO KNORR GALINHA 19g'},
  {id:57, cat:'MERCEARIA', nome:'Canela em pó 150g'},
  {id:58, cat:'MERCEARIA', nome:'Canjica BRANCA pacha 500g'},
  {id:59, cat:'MERCEARIA', nome:'Canjiquinha'},
  {id:60, cat:'MERCEARIA', nome:'CATCHUP SACHET predilecta'},
  {id:61, cat:'MERCEARIA', nome:'Chantilly spray Fleischmann'},
  {id:62, cat:'MERCEARIA', nome:'Cobertura calda chocolate e morango 1,3kg'},
  {id:63, cat:'MERCEARIA', nome:'Coco ralado'},
  {id:64, cat:'MERCEARIA', nome:'CREME DE LEITE 200 ML'},
  {id:65, cat:'MERCEARIA', nome:'Farinha de mandioca pacha 1 kg'},
  {id:66, cat:'MERCEARIA', nome:'FARINHA DE TRIGO GLOBO 5 KG'},
  {id:67, cat:'MERCEARIA', nome:'FARINHA DE TRIGO VENTURELLI 5 KG'},
  {id:68, cat:'MERCEARIA', nome:'Feijão codil 1 kg'},
  {id:69, cat:'MERCEARIA', nome:'FERMENTO PÓ ROYAL 250g'},
  {id:70, cat:'MERCEARIA', nome:'FÓSFORO PINHEIRO'},
  {id:71, cat:'MERCEARIA', nome:'Fubá SINHA'},
  {id:72, cat:'MERCEARIA', nome:'LEITE CONDENSADO 350g'},
  {id:73, cat:'MERCEARIA', nome:'Leite de Coco 500 ml'},
  {id:74, cat:'MERCEARIA', nome:'LEITE INTEGRAL ITAMBÉ 1 LITRO'},
  {id:75, cat:'MERCEARIA', nome:'Macarrão espaguete vilma 500g'},
  {id:76, cat:'MERCEARIA', nome:'MARGARINA PRIMOR 15 KG'},
  {id:77, cat:'MERCEARIA', nome:'Mel 1 litro'},
  {id:78, cat:'MERCEARIA', nome:'MILHO VERDE A VAPOR FUJINI 200g'},
  {id:79, cat:'MERCEARIA', nome:'MM\'S MAIS BARATO 1kg'},
  {id:80, cat:'MERCEARIA', nome:'MOLHO DE PIMENTA 900 ML Pirata'},
  {id:81, cat:'MERCEARIA', nome:'MOLHO DE PIMENTA extraforte 150 ML Pirata'},
  {id:82, cat:'MERCEARIA', nome:'Molho Shoyo Pirata 900 ml'},
  {id:83, cat:'MERCEARIA', nome:'MOLHO TOM. POMAROLA 340g'},
  {id:84, cat:'MERCEARIA', nome:'MOLHO TOM. SALSARETTI BOLONHESA 340g'},
  {id:85, cat:'MERCEARIA', nome:'ÓLEO DE SOJA 18 LITROS'},
  {id:86, cat:'MERCEARIA', nome:'ÓLEO PET SOJA 900 ML'},
  {id:87, cat:'MERCEARIA', nome:'ORÉGANO POTE 270g'},
  {id:88, cat:'MERCEARIA', nome:'Ovomaltine'},
  {id:89, cat:'MERCEARIA', nome:'OVOS'},
  {id:90, cat:'MERCEARIA', nome:'PALMITO SOBERANO PICADO 450g'},
  {id:91, cat:'MERCEARIA', nome:'PÃO DE QUEIJO CHEFF GOURMET'},
  {id:92, cat:'MERCEARIA', nome:'PURÊ DE BATATA ajinomoto 1kg'},
  {id:93, cat:'MERCEARIA', nome:'SAL EM SACHET'},
  {id:94, cat:'MERCEARIA', nome:'SAL GROSSO KG'},
  {id:95, cat:'MERCEARIA', nome:'SAL REFINADO KG'},
  {id:96, cat:'MERCEARIA', nome:'TRIGO P/ KIBE'},
  {id:97, cat:'MERCEARIA', nome:'VINAGRE BRANCO 750 ml'},
  {id:98, cat:'LIMPEZA', nome:'ÁGUA SANITÁRIA LITRO'},
  {id:99, cat:'LIMPEZA', nome:'ÁLCOOL LÍQ. 1 LITRO'},
  {id:100, cat:'LIMPEZA', nome:'BUCHA MULT USO'},
  {id:101, cat:'LIMPEZA', nome:'BUCHA PRETA de Fibra 125x88mm'},
  {id:102, cat:'LIMPEZA', nome:'DETERG. LÍQ. YPE/LIMPOL NEUTRO 500 ML'},
  {id:103, cat:'LIMPEZA', nome:'RODO'},
  {id:104, cat:'LIMPEZA', nome:'ROLO DE PANO MULT USO'},
  {id:105, cat:'LIMPEZA', nome:'SACO DE LIXO PRETO 150 LITROS'},
  {id:106, cat:'LIMPEZA', nome:'SACO DE LIXO PRETO 50 LITROS'},
  {id:107, cat:'LIMPEZA', nome:'VASSOURA'},
  {id:108, cat:'DESCARTÁVEIS', nome:'BARBANTE ROLO 310 METROS'},
  {id:109, cat:'DESCARTÁVEIS', nome:'BOBINA PAPEL MANTEIGA 60CM'},
  {id:110, cat:'DESCARTÁVEIS', nome:'BOBINA PERSONALIZADA PARA SALGADOS'},
  {id:111, cat:'DESCARTÁVEIS', nome:'BOBINA TERMICA 57MM'},
  {id:112, cat:'DESCARTÁVEIS', nome:'BOBINA TERMICA 79 MM'},
  {id:113, cat:'DESCARTÁVEIS', nome:'CANUDO EMPLASTIFICADO cx'},
  {id:114, cat:'DESCARTÁVEIS', nome:'COPO DESC. 10 ML'},
  {id:115, cat:'DESCARTÁVEIS', nome:'COPO DESC. 200 ML'},
  {id:116, cat:'DESCARTÁVEIS', nome:'COPO DESC. 300 ML'},
  {id:117, cat:'DESCARTÁVEIS', nome:'EMBALAGEM DE PIZZA 025'},
  {id:118, cat:'DESCARTÁVEIS', nome:'EMBALAGEM DE PIZZA 030'},
  {id:119, cat:'DESCARTÁVEIS', nome:'EMBALAGEM DE PIZZA 035'},
  {id:120, cat:'DESCARTÁVEIS', nome:'FITA ADESIVA ROLÃO'},
  {id:121, cat:'DESCARTÁVEIS', nome:'GARRAFA COM TAMPA DESC. 300 ML'},
  {id:122, cat:'DESCARTÁVEIS', nome:'GUARDANAPO PAPEL EM SACHET cx'},
  {id:123, cat:'DESCARTÁVEIS', nome:'MESINHA DE PIZZA'},
  {id:124, cat:'DESCARTÁVEIS', nome:'PALITOS DENTE EM SACHET'},
  {id:125, cat:'DESCARTÁVEIS', nome:'PAPEL HIGIÊNICO ROLÃO'},
  {id:126, cat:'DESCARTÁVEIS', nome:'PAPEL OFÍCIO A4'},
  {id:127, cat:'DESCARTÁVEIS', nome:'PAPEL TOALHA'},
  {id:128, cat:'DESCARTÁVEIS', nome:'POTE DE CALDO 500ML'},
  {id:129, cat:'DESCARTÁVEIS', nome:'PRATO LAMINADO PIZZA 025'},
  {id:130, cat:'DESCARTÁVEIS', nome:'PRATO LAMINADO PIZZA 030'},
  {id:131, cat:'DESCARTÁVEIS', nome:'PRATO LAMINADO PIZZA 035'},
  {id:132, cat:'DESCARTÁVEIS', nome:'PRATO PAPEL N-3'},
  {id:133, cat:'DESCARTÁVEIS', nome:'PRATO PAPEL N-4'},
  {id:134, cat:'DESCARTÁVEIS', nome:'PRATO PAPEL N-6'},
  {id:135, cat:'DESCARTÁVEIS', nome:'SACO EMBALAR 2 KG PIZZARIA GUARANI'},
  {id:136, cat:'DESCARTÁVEIS', nome:'SACOLA PERSONALIZADA PIZZARIA GUARANI'},
  {id:137, cat:'DESCARTÁVEIS', nome:'TOUCA DESCARTÁVEL'},
  {id:138, cat:'AÇOUGUE / FRIOS', nome:'APRESUNTADO kg'},
  {id:139, cat:'AÇOUGUE / FRIOS', nome:'BACON em cubos kg'},
  {id:140, cat:'AÇOUGUE / FRIOS', nome:'BACON em manta kg'},
  {id:141, cat:'AÇOUGUE / FRIOS', nome:'BATATA Frita DoCheff Congelada'},
  {id:142, cat:'AÇOUGUE / FRIOS', nome:'Cheddar carmolac'},
  {id:143, cat:'AÇOUGUE / FRIOS', nome:'Carne de sol kg'},
  {id:144, cat:'AÇOUGUE / FRIOS', nome:'CORAÇÃOZINHO kg'},
  {id:145, cat:'AÇOUGUE / FRIOS', nome:'FILÉ DE PEITO DE FRANGO kg'},
  {id:146, cat:'AÇOUGUE / FRIOS', nome:'FILÉ MIGNON kg'},
  {id:147, cat:'AÇOUGUE / FRIOS', nome:'LINGUIÇA CALABRESA'},
  {id:148, cat:'AÇOUGUE / FRIOS', nome:'LOMBO kg'},
  {id:149, cat:'AÇOUGUE / FRIOS', nome:'LOMBO CANADENSE seara kg'},
  {id:150, cat:'AÇOUGUE / FRIOS', nome:'Massa de Lasanha'},
  {id:151, cat:'AÇOUGUE / FRIOS', nome:'MUÇARELA kg'},
  {id:152, cat:'AÇOUGUE / FRIOS', nome:'Polpa acerola'},
  {id:153, cat:'AÇOUGUE / FRIOS', nome:'Polpa manga'},
  {id:154, cat:'AÇOUGUE / FRIOS', nome:'Polpa maracujá'},
  {id:155, cat:'AÇOUGUE / FRIOS', nome:'Polpa morango'},
  {id:156, cat:'AÇOUGUE / FRIOS', nome:'Provolone em barra grande 2irmãos'},
  {id:157, cat:'AÇOUGUE / FRIOS', nome:'Requeijão cremoso Cabeça Danta 1,8kg'},
  {id:158, cat:'AÇOUGUE / FRIOS', nome:'Tilápia 500g'},
  {id:159, cat:'AÇOUGUE / FRIOS', nome:'Torresmo kg'},
  {id:160, cat:'AÇOUGUE / FRIOS', nome:'Toucinho de barriga kg'},
  {id:161, cat:'SACOLÃO', nome:'Alho roxo N 6 especial'},
  {id:162, cat:'SACOLÃO', nome:'Cebola'},
  {id:163, cat:'SACOLÃO', nome:'Manjericão'},
  {id:164, cat:'SACOLÃO', nome:'Pimentão'},
  {id:165, cat:'SACOLÃO', nome:'Tomate andreia'},
  {id:166, cat:'SORVETE', nome:'Sorvete pote Chocolate'},
  {id:167, cat:'SORVETE', nome:'Sorvete pote Creme'},
  {id:168, cat:'SORVETE', nome:'Sorvete pote Flocos'},
  {id:169, cat:'SORVETE', nome:'Sorvete pote Morango'}
];

const MAX_FORNECEDORES = 10;
const ADMIN_SENHA = "guarani2026";
const SK_FORN  = "cotacao_fornecedores_v1";
const SK_PRECO = "cotacao_precos_v1";
const SK_QTD   = "cotacao_qtd_v1";

// Storage
const load = (k,fb) => { try { const v=localStorage.getItem(k); return v?JSON.parse(v):fb; } catch { return fb; } };
const save = (k,v)  => { try { localStorage.setItem(k,JSON.stringify(v)); } catch {} };

// Cores
const C = {
  bg:"#0a0d0f", card:"#141a1f", card2:"#1c2530",
  border:"#2a3540", text:"#e0eaf4", muted:"#6a8090",
  blue:"#1a6fa8", blueL:"#2a8fd8", green:"#1a7a3a", greenL:"#28a84e",
  gold:"#c8960a", goldL:"#f0b820", red:"#a82020", redL:"#d43030",
  white:"#ffffff",
};
const SI = (e={}) => ({
  width:"100%", padding:"8px 11px", borderRadius:7,
  border:`1.5px solid ${C.border}`, background:"#080c10",
  color:C.text, fontSize:13, fontFamily:"inherit",
  outline:"none", boxSizing:"border-box", ...e,
});
const SB = (bg,col="#fff",e={}) => ({
  padding:"9px 18px", borderRadius:8, border:"none",
  background:bg, color:col, fontWeight:700, fontSize:13,
  cursor:"pointer", fontFamily:"inherit", ...e,
});
const SC = (e={}) => ({
  background:C.card, borderRadius:12, border:`1px solid ${C.border}`,
  padding:"18px", ...e,
});

// ─── Tela inicial ─────────────────────────────────────────────────────────────
function TelaInicial({onAdmin, onFornecedor, fornecedores}) {
  const [mostrarForn, setMostrarForn] = useState(false);
  const [senha, setSenha] = useState("");
  const [erro, setErro] = useState("");

  const entrarAdmin = () => {
    if(senha === ADMIN_SENHA){ setErro(""); onAdmin(); }
    else setErro("Senha incorreta.");
  };

  const ativos = fornecedores.filter(f=>f.ativo);

  return (
    <div style={{minHeight:"100vh", background:C.bg, display:"flex", alignItems:"center",
      justifyContent:"center", padding:16, fontFamily:"'Segoe UI',sans-serif"}}>
      <div style={{width:"100%", maxWidth:440}}>
        {/* Header */}
        <div style={{textAlign:"center", marginBottom:32}}>
          <div style={{fontSize:40, marginBottom:8}}>🛒</div>
          <div style={{fontSize:22, fontWeight:800, color:C.text, letterSpacing:"-0.02em"}}>
            Sistema de Cotação
          </div>
          <div style={{fontSize:13, color:C.muted, marginTop:4}}>Pizzaria Guarani</div>
        </div>

        {/* Admin */}
        <div style={SC({marginBottom:14})}>
          <div style={{fontSize:13, fontWeight:700, color:C.goldL, marginBottom:12}}>
            🔐 Acesso Administrador
          </div>
          <input style={SI({marginBottom:10})} type="password" placeholder="Senha do administrador"
            value={senha} onChange={e=>setSenha(e.target.value)}
            onKeyDown={e=>e.key==="Enter"&&entrarAdmin()}/>
          {erro && <div style={{color:C.redL, fontSize:12, marginBottom:8}}>{erro}</div>}
          <button onClick={entrarAdmin} style={{...SB(C.gold, "#000"), width:"100%", padding:"10px"}}>
            Entrar como Admin
          </button>
        </div>

        {/* Fornecedores */}
        <div style={SC()}>
          <div style={{fontSize:13, fontWeight:700, color:C.blueL, marginBottom:12}}>
            🏪 Acesso Fornecedor
          </div>
          {ativos.length === 0 ? (
            <div style={{fontSize:13, color:C.muted, textAlign:"center", padding:"12px 0"}}>
              Nenhum fornecedor cadastrado ainda. O admin deve cadastrar primeiro.
            </div>
          ) : (
            <>
              {!mostrarForn ? (
                <button onClick={()=>setMostrarForn(true)}
                  style={{...SB(C.blue), width:"100%", padding:"10px"}}>
                  Selecionar meu fornecedor
                </button>
              ) : (
                <div style={{display:"flex", flexDirection:"column", gap:8}}>
                  {ativos.map(f => (
                    <button key={f.id} onClick={()=>onFornecedor(f)}
                      style={{...SB(C.card2, C.text), width:"100%", padding:"10px 14px",
                        textAlign:"left", border:`1px solid ${C.border}`,
                        display:"flex", alignItems:"center", gap:10}}>
                      <span style={{fontSize:20}}>🏪</span>
                      <div>
                        <div style={{fontWeight:700, fontSize:14}}>{f.nome}</div>
                        <div style={{fontSize:11, color:C.muted}}>{f.contato||"—"}</div>
                      </div>
                    </button>
                  ))}
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── Tela Fornecedor ──────────────────────────────────────────────────────────
function TelaFornecedor({fornecedor, precos, onSalvar, onVoltar}) {
  const [local, setLocal] = useState(() => {
    const init = {};
    ITEMS.forEach(it => { init[it.id] = precos[fornecedor.id]?.[it.id] ?? ""; });
    return init;
  });
  const [salvando, setSalvando] = useState(false);
  const [salvo, setSalvo] = useState(false);
  const [filtro, setFiltro] = useState("TODAS");

  const categorias = ["TODAS", ...new Set(ITEMS.map(i=>i.cat))];
  const itensFiltrados = filtro === "TODAS" ? ITEMS : ITEMS.filter(i=>i.cat===filtro);

  const set = (id, val) => setLocal(p=>({...p,[id]:val}));

  const salvar = () => {
    setSalvando(true);
    setTimeout(()=>{
      onSalvar(fornecedor.id, local);
      setSalvando(false); setSalvo(true);
      setTimeout(()=>setSalvo(false), 3000);
    }, 500);
  };

  const preenchidos = Object.values(local).filter(v=>v!=="").length;

  return (
    <div style={{minHeight:"100vh", background:C.bg, fontFamily:"'Segoe UI',sans-serif"}}>
      {/* Header */}
      <div style={{background:C.card, borderBottom:`1px solid ${C.border}`,
        padding:"12px 18px", display:"flex", alignItems:"center",
        justifyContent:"space-between", gap:10, position:"sticky", top:0, zIndex:10}}>
        <div>
          <div style={{fontSize:15, fontWeight:700, color:C.text}}>🛒 Cotação de Preços</div>
          <div style={{fontSize:12, color:C.blueL, fontWeight:600}}>{fornecedor.nome}</div>
        </div>
        <div style={{display:"flex", alignItems:"center", gap:10}}>
          <div style={{fontSize:12, color:C.muted}}>
            <span style={{color:C.goldL, fontWeight:700}}>{preenchidos}</span>/{ITEMS.length} itens
          </div>
          <button onClick={onVoltar} style={{...SB(C.card2, C.muted), padding:"6px 12px", fontSize:12}}>
            Sair
          </button>
        </div>
      </div>

      <div style={{maxWidth:700, margin:"0 auto", padding:"16px 12px 80px"}}>
        {salvo && (
          <div style={{background:"#0d2e18", border:`1px solid ${C.greenL}55`, borderRadius:10,
            padding:"12px 15px", marginBottom:14, color:C.greenL, fontWeight:600}}>
            ✅ Preços salvos com sucesso!
          </div>
        )}

        {/* Instrução */}
        <div style={SC({marginBottom:14, background:"#0d1520", border:`1px solid ${C.blue}33`})}>
          <div style={{fontSize:13, color:C.muted, lineHeight:1.7}}>
            📋 Preencha o <strong style={{color:C.text}}>preço unitário</strong> de cada item que você fornece.
            Deixe em branco os itens que não trabalha. Os preços são <strong style={{color:C.goldL}}>confidenciais</strong> — somente o administrador pode ver e comparar.
          </div>
        </div>

        {/* Filtro por categoria */}
        <div style={{display:"flex", gap:6, flexWrap:"wrap", marginBottom:16}}>
          {categorias.map(cat => (
            <button key={cat} onClick={()=>setFiltro(cat)}
              style={{...SB(filtro===cat?C.blue:C.card2, filtro===cat?C.white:C.muted),
                padding:"5px 12px", fontSize:11, borderRadius:20,
                border:`1px solid ${filtro===cat?C.blue:C.border}`}}>
              {cat}
            </button>
          ))}
        </div>

        {/* Lista de itens agrupada */}
        {(filtro==="TODAS" ? [...new Set(ITEMS.map(i=>i.cat))] : [filtro]).map(cat => (
          <div key={cat} style={{marginBottom:20}}>
            <div style={{fontSize:11, fontWeight:700, color:C.blueL, textTransform:"uppercase",
              letterSpacing:"0.08em", marginBottom:8, display:"flex", alignItems:"center", gap:8}}>
              <span style={{width:8, height:8, borderRadius:"50%", background:C.blueL, display:"inline-block"}}/>
              {cat}
            </div>
            <div style={SC({padding:"0"})}>
              {ITEMS.filter(i=>i.cat===cat).map((item, idx, arr) => (
                <div key={item.id} style={{
                  display:"grid", gridTemplateColumns:"1fr 160px",
                  alignItems:"center", gap:12,
                  padding:"10px 16px",
                  borderBottom: idx<arr.length-1 ? `1px solid ${C.border}` : "none",
                  background: idx%2===0 ? "transparent" : "#0d1218",
                }}>
                  <div style={{fontSize:13, color:C.text}}>{item.nome}</div>
                  <div style={{position:"relative"}}>
                    <span style={{position:"absolute", left:10, top:"50%",
                      transform:"translateY(-50%)", color:C.muted, fontSize:13, fontWeight:700}}>
                      R$
                    </span>
                    <input
                      style={{...SI({paddingLeft:32, textAlign:"right"}),
                        borderColor: local[item.id] ? C.blueL : C.border}}
                      type="number" step="0.01" min="0" placeholder="0,00"
                      value={local[item.id]}
                      onChange={e=>set(item.id, e.target.value)}
                    />
                  </div>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>

      {/* Footer fixo */}
      <div style={{position:"fixed", bottom:0, left:0, right:0,
        background:C.card, borderTop:`1px solid ${C.border}`, padding:"12px 18px",
        display:"flex", gap:12, justifyContent:"center"}}>
        <button onClick={salvar} disabled={salvando}
          style={{...SB(C.green), padding:"12px 40px", fontSize:15,
            opacity:salvando?0.7:1, boxShadow:`0 4px 16px ${C.green}44`}}>
          {salvando ? "Salvando..." : "💾 Salvar Cotação"}
        </button>
      </div>
    </div>
  );
}

// ─── Painel Admin ─────────────────────────────────────────────────────────────
function PainelAdmin({fornecedores, setFornecedores, precos, qtds, setQtds, onVoltar}) {
  const [aba, setAba] = useState("comparativo");
  const [modalForn, setModalForn] = useState(null);
  const [formForn, setFormForn] = useState({nome:"", contato:""});
  const [erroForn, setErroForn] = useState("");
  const [filtCat, setFiltCat] = useState("TODAS");
  const [filtForn, setFiltForn] = useState("TODOS");

  const categorias = ["TODAS", ...new Set(ITEMS.map(i=>i.cat))];
  const fornAtivos = fornecedores.filter(f=>f.ativo);

  // Calcular melhor preço por item
  const analise = useMemo(() => {
    return ITEMS.map(item => {
      const ofertas = fornAtivos
        .map(f => ({ forn: f, preco: parseFloat(precos[f.id]?.[item.id]) || null }))
        .filter(o => o.preco !== null && o.preco > 0)
        .sort((a,b) => a.preco - b.preco);
      const melhor = ofertas[0] || null;
      return { item, ofertas, melhor };
    });
  }, [fornecedores, precos]);

  // Resumo por fornecedor (onde comprar)
  const resumoPorForn = useMemo(() => {
    const mapa = {};
    fornAtivos.forEach(f => { mapa[f.id] = {forn:f, itens:[], total:0}; });
    analise.forEach(({item, melhor, ofertas}) => {
      if(!melhor) return;
      const qtd = parseFloat(qtds[item.id]) || 0;
      mapa[melhor.forn.id]?.itens.push({
        item, preco:melhor.preco,
        economia: ofertas.length>1 ? (ofertas[1].preco - melhor.preco) : 0,
        subtotal: qtd * melhor.preco,
      });
      if(mapa[melhor.forn.id]) mapa[melhor.forn.id].total += qtd * melhor.preco;
    });
    return Object.values(mapa).filter(v=>v.itens.length>0).sort((a,b)=>b.itens.length-a.itens.length);
  }, [analise, qtds]);

  // Salvar fornecedor
  const salvarForn = () => {
    if(!formForn.nome.trim()){setErroForn("Nome obrigatório.");return;}
    const dup = fornecedores.find(f=>f.nome.toLowerCase()===formForn.nome.toLowerCase()&&
      (modalForn.modo==="novo"||f.id!==modalForn.id));
    if(dup){setErroForn("Já existe fornecedor com este nome.");return;}
    if(modalForn.modo==="novo"){
      const nid = fornecedores.length>0 ? Math.max(...fornecedores.map(f=>f.id))+1 : 1;
      setFornecedores(p=>[...p,{id:nid,...formForn,ativo:true}]);
    } else {
      setFornecedores(p=>p.map(f=>f.id===modalForn.id?{...f,...formForn}:f));
    }
    setModalForn(null);
  };

  // Exportar Excel
  const exportarExcel = () => {
    const wb = XLSX.utils.book_new();

    // Aba 1 — Comparativo geral
    const header = ["Categoria","Item","Qtd Compra",...fornAtivos.map(f=>f.nome),"Menor Preço","Fornecedor","Subtotal (R$)"];
    const rows = [header];
    const catFilt = filtCat==="TODAS" ? [...new Set(ITEMS.map(i=>i.cat))] : [filtCat];
    catFilt.forEach(cat => {
      analise.filter(a=>a.item.cat===cat).forEach(({item,ofertas,melhor}) => {
        const qtd = parseFloat(qtds[item.id])||0;
        const precosRow = fornAtivos.map(f => {
          const p = parseFloat(precos[f.id]?.[item.id])||null;
          return p ? p : "—";
        });
        rows.push([
          item.cat, item.nome, qtd||"",
          ...precosRow,
          melhor ? melhor.preco : "—",
          melhor ? melhor.forn.nome : "—",
          melhor && qtd ? (melhor.preco*qtd).toFixed(2) : "—",
        ]);
      });
    });
    const ws1 = XLSX.utils.aoa_to_sheet(rows);
    ws1["!cols"] = [14,35,10,...fornAtivos.map(()=>14),14,20,14].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws1, "Comparativo Geral");

    // Aba 2 — Lista de compras por fornecedor
    resumoPorForn.forEach(({forn, itens, total}) => {
      const r = [["Item","Categoria","Preço Unit.","Qtd","Subtotal","Economia vs 2º"]];
      itens.forEach(({item,preco,subtotal,economia}) => {
        const qtd = parseFloat(qtds[item.id])||0;
        r.push([item.nome,item.cat,preco,qtd||"",subtotal?subtotal.toFixed(2):"—",
          economia>0?economia.toFixed(2):"—"]);
      });
      r.push(["","","","TOTAL",total>0?total.toFixed(2):"—",""]);
      const ws = XLSX.utils.aoa_to_sheet(r);
      ws["!cols"] = [35,16,12,10,12,14].map(w=>({wch:w}));
      XLSX.utils.book_append_sheet(wb, ws, forn.nome.substring(0,30));
    });

    XLSX.writeFile(wb, `cotacao_guarani_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  const itensFiltrados = analise.filter(a =>
    (filtCat==="TODAS"||a.item.cat===filtCat) &&
    (filtForn==="TODOS"||!a.melhor||a.melhor.forn.nome===filtForn)
  );

  return (
    <div style={{minHeight:"100vh", background:C.bg, fontFamily:"'Segoe UI',sans-serif"}}>
      {/* Modal fornecedor */}
      {modalForn && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.85)",display:"flex",
          alignItems:"center",justifyContent:"center",zIndex:999,padding:16}}>
          <div style={{background:C.card,borderRadius:14,border:`1px solid ${C.border}`,
            padding:"24px 22px",width:"100%",maxWidth:380}}>
            <div style={{fontSize:16,fontWeight:800,color:C.text,marginBottom:18}}>
              {modalForn.modo==="novo"?"➕ Novo Fornecedor":"✏️ Editar Fornecedor"}
            </div>
            {[{k:"nome",l:"Nome do fornecedor"},{k:"contato",l:"Contato (opcional)"}].map(({k,l})=>(
              <div key={k} style={{marginBottom:12}}>
                <label style={{display:"block",fontSize:11,fontWeight:700,color:C.muted,
                  textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:4}}>{l}</label>
                <input style={SI()} type="text" value={formForn[k]}
                  onChange={e=>setFormForn(p=>({...p,[k]:e.target.value}))}/>
              </div>
            ))}
            {erroForn&&<div style={{color:C.redL,fontSize:12,marginBottom:10}}>{erroForn}</div>}
            <div style={{display:"flex",gap:10,marginTop:4}}>
              <button onClick={()=>setModalForn(null)} style={{...SB(C.card2,C.muted),flex:1}}>Cancelar</button>
              <button onClick={salvarForn} style={{...SB(C.blue),flex:2}}>
                {modalForn.modo==="novo"?"Criar":"Salvar"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Header */}
      <div style={{background:C.card,borderBottom:`1px solid ${C.border}`,padding:"12px 18px",
        display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
        <div>
          <div style={{fontSize:15,fontWeight:700,color:C.text}}>🛒 Painel Admin — Cotação</div>
          <div style={{fontSize:12,color:C.muted}}>Pizzaria Guarani</div>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={exportarExcel}
            style={{...SB(C.green),padding:"8px 14px",fontSize:12}}>
            📊 Exportar Excel
          </button>
          <button onClick={onVoltar}
            style={{...SB(C.card2,C.muted),padding:"8px 12px",fontSize:12}}>
            Sair
          </button>
        </div>
      </div>

      {/* Abas */}
      <div style={{background:C.card,borderBottom:`1px solid ${C.border}`,display:"flex",overflowX:"auto"}}>
        {[["comparativo","📋 Comparativo"],["fornecedores","🏪 Fornecedores"],["onde_comprar","🏆 Onde Comprar"]].map(([a,l])=>(
          <button key={a} onClick={()=>setAba(a)} style={{padding:"11px 18px",border:"none",
            background:"transparent",fontWeight:700,fontSize:13,cursor:"pointer",whiteSpace:"nowrap",
            color:aba===a?C.goldL:C.muted,
            borderBottom:aba===a?`2px solid ${C.goldL}`:"2px solid transparent"}}>
            {l}
          </button>
        ))}
      </div>

      <div style={{maxWidth:1100,margin:"0 auto",padding:"16px 12px 60px"}}>

        {/* ── COMPARATIVO ── */}
        {aba==="comparativo"&&(
          <>
            {/* KPIs */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:10,marginBottom:16}}>
              {[
                {l:"Fornecedores",v:fornAtivos.length,c:C.blueL},
                {l:"Itens cotados",v:analise.filter(a=>a.ofertas.length>0).length,c:C.goldL},
                {l:"Sem cotação",v:analise.filter(a=>a.ofertas.length===0).length,c:C.muted},
                {l:"Total estimado",v:`R$ ${analise.reduce((acc,{melhor})=>{
                  if(!melhor)return acc;
                  return acc+(melhor.preco*(parseFloat(qtds[melhor.item?.id||0])||0));
                },0).toFixed(2)}`,c:C.greenL},
              ].map(k=>(
                <div key={k.l} style={SC({textAlign:"center",padding:"14px 10px"})}>
                  <div style={{fontSize:20,fontWeight:800,color:k.c}}>{k.v}</div>
                  <div style={{fontSize:11,color:C.muted,marginTop:3}}>{k.l}</div>
                </div>
              ))}
            </div>

            {/* Filtros */}
            <div style={SC({marginBottom:14,display:"flex",flexWrap:"wrap",gap:10})}>
              <div style={{flex:1,minWidth:160}}>
                <label style={{display:"block",fontSize:11,fontWeight:700,color:C.muted,marginBottom:4,textTransform:"uppercase"}}>Categoria</label>
                <select style={SI()} value={filtCat} onChange={e=>setFiltCat(e.target.value)}>
                  {categorias.map(c=><option key={c}>{c}</option>)}
                </select>
              </div>
              <div style={{flex:1,minWidth:160}}>
                <label style={{display:"block",fontSize:11,fontWeight:700,color:C.muted,marginBottom:4,textTransform:"uppercase"}}>Melhor preço por</label>
                <select style={SI()} value={filtForn} onChange={e=>setFiltForn(e.target.value)}>
                  <option value="TODOS">Todos</option>
                  {fornAtivos.map(f=><option key={f.id}>{f.nome}</option>)}
                </select>
              </div>
            </div>

            {/* Tabela comparativa */}
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead>
                  <tr style={{background:C.card2}}>
                    <th style={{padding:"10px 12px",textAlign:"left",color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase",whiteSpace:"nowrap",borderBottom:`1px solid ${C.border}`}}>Item</th>
                    <th style={{padding:"10px 8px",textAlign:"center",color:C.muted,fontSize:11,fontWeight:700,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>Qtd Compra</th>
                    {fornAtivos.map(f=>(
                      <th key={f.id} style={{padding:"10px 8px",textAlign:"center",color:C.blueL,fontSize:11,fontWeight:700,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>{f.nome}</th>
                    ))}
                    <th style={{padding:"10px 8px",textAlign:"center",color:C.goldL,fontSize:11,fontWeight:700,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>✓ Menor</th>
                    <th style={{padding:"10px 8px",textAlign:"center",color:C.greenL,fontSize:11,fontWeight:700,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>Subtotal</th>
                  </tr>
                </thead>
                <tbody>
                  {itensFiltrados.map(({item,ofertas,melhor},idx)=>{
                    const qtd = parseFloat(qtds[item.id])||0;
                    return(
                      <tr key={item.id} style={{background:idx%2===0?C.card:"#0e1520",
                        borderBottom:`1px solid ${C.border}33`}}>
                        <td style={{padding:"9px 12px",color:C.text,fontWeight:500}}>
                          <div style={{fontSize:12,color:C.muted,marginBottom:1}}>{item.cat}</div>
                          {item.nome}
                        </td>
                        <td style={{padding:"6px 8px"}}>
                          <input
                            style={{...SI({width:80,textAlign:"center",padding:"5px 6px",fontSize:12}),margin:"0 auto",display:"block"}}
                            type="number" min="0" step="1" placeholder="0"
                            value={qtds[item.id]||""}
                            onChange={e=>setQtds(p=>({...p,[item.id]:e.target.value}))}
                          />
                        </td>
                        {fornAtivos.map(f=>{
                          const p = parseFloat(precos[f.id]?.[item.id])||null;
                          const isMelhor = melhor && melhor.forn.id===f.id;
                          return(
                            <td key={f.id} style={{padding:"9px 8px",textAlign:"center",
                              background:isMelhor?"#0a2010":"transparent",
                              color:isMelhor?C.greenL:p?C.text:C.muted,
                              fontWeight:isMelhor?700:400}}>
                              {p ? `R$ ${p.toFixed(2)}` : "—"}
                              {isMelhor && <span style={{marginLeft:4}}>✓</span>}
                            </td>
                          );
                        })}
                        <td style={{padding:"9px 8px",textAlign:"center",color:C.goldL,fontWeight:700}}>
                          {melhor ? `R$ ${melhor.preco.toFixed(2)}` : "—"}
                          {melhor && <div style={{fontSize:10,color:C.muted}}>{melhor.forn.nome}</div>}
                        </td>
                        <td style={{padding:"9px 8px",textAlign:"center",
                          color:melhor&&qtd?C.greenL:C.muted,fontWeight:700}}>
                          {melhor&&qtd ? `R$ ${(melhor.preco*qtd).toFixed(2)}` : "—"}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </>
        )}

        {/* ── ONDE COMPRAR ── */}
        {aba==="onde_comprar"&&(
          <>
            <div style={{fontSize:13,color:C.muted,marginBottom:16}}>
              Resumo dos melhores preços agrupados por fornecedor. Preencha as quantidades na aba Comparativo.
            </div>
            {resumoPorForn.length===0?(
              <div style={SC({textAlign:"center",padding:"40px",color:C.muted})}>
                Nenhum preço cotado ainda.
              </div>
            ):(
              resumoPorForn.map(({forn,itens,total})=>(
                <div key={forn.id} style={SC({marginBottom:14})}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",
                    flexWrap:"wrap",gap:8,marginBottom:14}}>
                    <div style={{display:"flex",alignItems:"center",gap:10}}>
                      <span style={{fontSize:22}}>🏪</span>
                      <div>
                        <div style={{fontSize:16,fontWeight:700,color:C.text}}>{forn.nome}</div>
                        <div style={{fontSize:12,color:C.muted}}>{itens.length} itens com menor preço</div>
                      </div>
                    </div>
                    <div style={{textAlign:"right"}}>
                      <div style={{fontSize:11,color:C.muted}}>Total estimado</div>
                      <div style={{fontSize:18,fontWeight:800,color:C.greenL}}>
                        R$ {total>0?total.toFixed(2):"—"}
                      </div>
                    </div>
                  </div>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                    <thead>
                      <tr>
                        {["Item","Categoria","Preço Unit.","Qtd","Subtotal"].map(h=>(
                          <th key={h} style={{padding:"7px 10px",textAlign:"left",
                            color:C.muted,fontSize:11,fontWeight:700,textTransform:"uppercase",
                            borderBottom:`1px solid ${C.border}`}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {itens.map(({item,preco,subtotal},i)=>{
                        const qtd = parseFloat(qtds[item.id])||0;
                        return(
                          <tr key={item.id} style={{background:i%2===0?"transparent":"#0d1218",
                            borderBottom:`1px solid ${C.border}22`}}>
                            <td style={{padding:"8px 10px",color:C.text}}>{item.nome}</td>
                            <td style={{padding:"8px 10px",color:C.muted,fontSize:12}}>{item.cat}</td>
                            <td style={{padding:"8px 10px",color:C.goldL,fontWeight:700}}>R$ {preco.toFixed(2)}</td>
                            <td style={{padding:"8px 10px",color:C.text}}>{qtd||"—"}</td>
                            <td style={{padding:"8px 10px",color:qtd?C.greenL:C.muted,fontWeight:700}}>
                              {qtd ? `R$ ${subtotal.toFixed(2)}` : "—"}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              ))
            )}
          </>
        )}

        {/* ── FORNECEDORES ── */}
        {aba==="fornecedores"&&(
          <>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
              <div style={{fontSize:13,color:C.muted}}>
                {fornAtivos.length}/{MAX_FORNECEDORES} fornecedores ativos
              </div>
              {fornAtivos.length<MAX_FORNECEDORES&&(
                <button onClick={()=>{setFormForn({nome:"",contato:""});setErroForn("");setModalForn({modo:"novo"})}}
                  style={{...SB(C.blue),padding:"9px 16px",fontSize:13}}>
                  ➕ Novo Fornecedor
                </button>
              )}
            </div>

            <div style={SC({marginBottom:16,background:"#0a1520",border:`1px solid ${C.blue}33`})}>
              <div style={{fontSize:13,fontWeight:700,color:C.blueL,marginBottom:6}}>
                📌 Como enviar para os fornecedores
              </div>
              <div style={{fontSize:13,color:C.muted,lineHeight:1.8}}>
                1. Cadastre cada fornecedor aqui<br/>
                2. Compartilhe o link do sistema pelo WhatsApp<br/>
                3. O fornecedor abre, clica no nome dele e preenche os preços<br/>
                4. <strong style={{color:C.text}}>Nenhum fornecedor vê o preço do outro</strong> — só o admin vê tudo
              </div>
            </div>

            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {fornecedores.map(f=>{
                const qtdCotados = ITEMS.filter(i=>parseFloat(precos[f.id]?.[i.id])>0).length;
                return(
                  <div key={f.id} style={SC({padding:"13px 16px",display:"flex",
                    alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10,
                    opacity:f.ativo?1:0.5})}>
                    <div style={{display:"flex",alignItems:"center",gap:12}}>
                      <span style={{fontSize:24}}>🏪</span>
                      <div>
                        <div style={{fontSize:14,fontWeight:700,color:C.text}}>{f.nome}</div>
                        <div style={{fontSize:12,color:C.muted}}>
                          {f.contato||"—"} · <span style={{color:C.goldL}}>{qtdCotados} itens cotados</span>
                          {!f.ativo&&<span style={{color:C.redL,marginLeft:8}}>● Inativo</span>}
                        </div>
                      </div>
                    </div>
                    <div style={{display:"flex",gap:8}}>
                      <button onClick={()=>{setFormForn({nome:f.nome,contato:f.contato||""});setErroForn("");setModalForn({modo:"editar",id:f.id})}}
                        style={{...SB(C.card2,C.blueL),padding:"7px 12px",fontSize:12}}>✏️ Editar</button>
                      <button onClick={()=>setFornecedores(p=>p.map(x=>x.id===f.id?{...x,ativo:!x.ativo}:x))}
                        style={{...SB(C.card2,f.ativo?C.redL:C.greenL),padding:"7px 12px",fontSize:12}}>
                        {f.ativo?"🚫 Desativar":"✅ Ativar"}
                      </button>
                    </div>
                  </div>
                );
              })}
              {fornecedores.length===0&&(
                <div style={SC({textAlign:"center",padding:"40px",color:C.muted})}>
                  Nenhum fornecedor cadastrado. Clique em "Novo Fornecedor".
                </div>
              )}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ─── App Root ─────────────────────────────────────────────────────────────────
export default function App() {
  const [fornecedores, setFornRaw]   = useState(()=>load(SK_FORN, []));
  const [precos,       setPrecosRaw] = useState(()=>load(SK_PRECO, {}));
  const [qtds,         setQtdsRaw]   = useState(()=>load(SK_QTD, {}));
  const [tela, setTela] = useState("inicio"); // inicio | admin | fornecedor
  const [fornAtual, setFornAtual] = useState(null);

  const setFornecedores = v => { const nv=typeof v==="function"?v(fornecedores):v; setFornRaw(nv); save(SK_FORN,nv); };
  const setPrecos = v => { const nv=typeof v==="function"?v(precos):v; setPrecosRaw(nv); save(SK_PRECO,nv); };
  const setQtds   = v => { const nv=typeof v==="function"?v(qtds):v;   setQtdsRaw(nv);   save(SK_QTD,nv);  };

  const salvarPrecosForn = (fornId, precosForn) => {
    setPrecos(p=>({...p, [fornId]: precosForn}));
  };

  if(tela==="admin") return (
    <PainelAdmin
      fornecedores={fornecedores} setFornecedores={setFornecedores}
      precos={precos} qtds={qtds} setQtds={setQtds}
      onVoltar={()=>setTela("inicio")}
    />
  );

  if(tela==="fornecedor" && fornAtual) return (
    <TelaFornecedor
      fornecedor={fornAtual} precos={precos}
      onSalvar={salvarPrecosForn}
      onVoltar={()=>{ setTela("inicio"); setFornAtual(null); }}
    />
  );

  return (
    <TelaInicial
      fornecedores={fornecedores}
      onAdmin={()=>setTela("admin")}
      onFornecedor={f=>{ setFornAtual(f); setTela("fornecedor"); }}
    />
  );
}
