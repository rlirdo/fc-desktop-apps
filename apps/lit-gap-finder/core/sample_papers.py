"""離線自測樣本（selftest/samples.json 的內建備援）。

書目資料取自 Crossref／doi.org 查證結果（見 9 篇文獻出處真實性查證表_APA7.xlsx 與 Crossref API），
非手寫杜撰；source=manual 者為「手動加入」示範，匯出時會附註「使用者自填，未經 API 驗證」。
cited_by／abstract 為離線示意值（近似），連網時以 API 實值覆蓋。
"""


def _a(*names):
    out = []
    for n in names:
        fam, _, giv = n.partition("|")
        out.append({"family": fam, "given": giv})
    return out


SAMPLES = [
    dict(doi="10.1515/plm-2012-0034", title="Far infrared radiation (FIR): Its biological effects and medical applications",
         authors=_a("Vatansever|Fatma", "Hamblin|Michael R."), year=2012, container="Photonics & Lasers in Medicine",
         volume="1", issue="4", pages="255-266", type="journal-article", issn=["2193-0643", "2193-0635"],
         publisher="Walter de Gruyter GmbH", language="en", cited_by=300, source="crossref",
         abstract="Far infrared (FIR) radiation biological effects and medical applications review."),
    dict(doi="", title="遠紅外線散熱陶瓷漿料、使用該漿料製備之纖維布、使用該漿料製備之薄膜及其製備方法",
         authors=_a("吳秋香|", "謝建俊|"), year=2014, container="", type="patent", patent_number="TW201402514A",
         patent_office="中華民國專利公開", publisher="經濟部智慧財產局", language="zh", source="manual",
         note="使用者自填，未經 API 驗證"),
    dict(doi="10.1021/acscentsci.6b00277", title="Mechanochemistry: A force of synthesis",
         authors=_a("Do|Jean-Louis", "Friščić|Tomislav"), year=2017, container="ACS Central Science",
         volume="3", issue="1", pages="13-19", type="journal-article", issn=["2374-7943", "2374-7951"],
         publisher="American Chemical Society (ACS)", language="en", cited_by=900, oa_status="gold", source="crossref",
         abstract="Mechanochemistry, reactions driven by mechanical force such as ball milling, for green synthesis."),
    dict(doi="10.1016/j.heliyon.2018.e00923", title="Development of an electrolysis based system to continuously recover magnesium from seawater",
         authors=_a("Sano|Yoshihiko", "Hao|Yuxin", "Kuwahara|Fujio"), year=2018, container="Heliyon",
         volume="4", issue="11", article_number="e00923", type="journal-article", issn=["2405-8440"],
         publisher="Elsevier BV", language="en", cited_by=40, oa_status="gold", source="crossref",
         abstract="Electrolysis based system to recover magnesium from seawater."),
    dict(doi="10.1039/d0sc06000c", title="Water as the reaction medium in organic chemistry: From our worst enemy to our best friend",
         authors=_a("Cortes-Clerget|Margery", "Yu|Julie", "Kincaid|Joseph R. A.", "Walde|Peter", "Gallou|Fabrice", "Lipshutz|Bruce H."),
         year=2021, container="Chemical Science", volume="12", issue="12", pages="4237-4266", type="journal-article",
         issn=["2041-6520", "2041-6539"], publisher="Royal Society of Chemistry (RSC)", language="en", cited_by=350,
         oa_status="gold", source="crossref",
         abstract="Water as reaction medium with surfactant micelle technology such as TPGS-750-M."),
    dict(doi="10.3390/mti5070037", title="An overview of twenty-five years of augmented reality in education",
         authors=_a("Garzón|Juan"), year=2021, container="Multimodal Technologies and Interaction", volume="5", issue="7",
         article_number="37", type="journal-article", issn=["2414-4088"], publisher="MDPI AG", language="en",
         cited_by=200, oa_status="gold", source="crossref",
         abstract="Twenty-five years of augmented reality in education research reviewed."),
    dict(doi="10.3390/ma15248785", title="Magnesium silicate binding materials formed from heat-treated serpentine-group minerals and aqueous solutions: Structural features, acid-neutralizing capacity, and strength properties",
         authors=_a("Ivanova|Tatiana K.", "Kremenetskaya|Irina P.", "Marchevskaya|Vera V.", "Slukovskaya|Marina V.", "Drogobuzhskaya|Svetlana V."),
         year=2022, container="Materials", volume="15", issue="24", article_number="8785", type="journal-article",
         issn=["1996-1944"], publisher="MDPI AG", language="en", cited_by=8, oa_status="gold", source="crossref",
         abstract="Heat-treated serpentine minerals form magnesium silicate binding materials."),
    dict(doi="10.3390/min13030318", title="Recovery of Mg from H2SO4 leaching solution of serpentine to precipitation of high-purity Mg(OH)2 and 4MgCO3·Mg(OH)2·4H2O",
         authors=_a("Chen|Yang", "Yang|Xue", "Wu|Lin", "Tong|Ling", "Zhu|Jie"), year=2023, container="Minerals", volume="13",
         issue="3", article_number="318", type="journal-article", issn=["2075-163X"], publisher="MDPI AG", language="en",
         cited_by=10, oa_status="gold", source="crossref",
         abstract="Magnesium recovered from sulfuric acid leaching solution of serpentine."),
    dict(doi="10.3390/min15010062", title="Serpentinite applications: Effects of surface-ions-modified natural silicate minerals on cultivation of magnesium–manganese-enriched garlics",
         authors=_a("Hung|Fei Shan"), year=2025, container="Minerals", volume="15", issue="1", article_number="62",
         type="journal-article", issn=["2075-163X"], publisher="MDPI AG", language="en", cited_by=1, oa_status="gold",
         source="crossref", abstract="Surface-ions-modified serpentinite minerals applied to garlic cultivation."),
    dict(doi="10.1016/0749-5978(91)90020-t", title="The theory of planned behavior", authors=_a("Ajzen|Icek"), year=1991,
         container="Organizational Behavior and Human Decision Processes", volume="50", issue="2", pages="179-211",
         type="journal-article", issn=["0749-5978"], publisher="Elsevier BV", language="en", cited_by=90000, source="crossref",
         abstract="Theory of planned behavior: attitudes, subjective norms and perceived behavioral control."),
    dict(doi="10.1021/jo101974u", title="TPGS-750-M: A second-generation amphiphile for metal-catalyzed cross-couplings in water at room temperature",
         authors=_a("Lipshutz|Bruce H.", "Ghorai|Subir", "Abela|Alexander R.", "Moser|Ralph", "Nishikata|Takashi", "Duplais|Christophe",
                    "Krasovskiy|Arkady", "Gaston|Ricky D.", "Gadwood|Robert C."),
         year=2011, container="The Journal of Organic Chemistry", volume="76", issue="11", pages="4379-4391", type="journal-article",
         issn=["0022-3263", "1520-6904"], publisher="American Chemical Society (ACS)", language="en", cited_by=600,
         oa_status="green", source="crossref",
         abstract="TPGS-750-M designer surfactant micelles enable cross-couplings in water."),
    dict(doi="10.1039/c5gc02371h", title="Surfactant technology applied toward an active pharmaceutical ingredient: More than a simple green chemistry advance",
         authors=_a("Gallou|Fabrice", "Isley|Nicholas A.", "Ganic|Adnan", "Onken|Ulrich", "Parmentier|Michael"), year=2016,
         container="Green Chemistry", volume="18", issue="1", pages="14-19", type="journal-article", issn=["1463-9262", "1463-9270"],
         publisher="Royal Society of Chemistry (RSC)", language="en", cited_by=200, source="crossref",
         abstract="Surfactant micelle technology applied to active pharmaceutical ingredient synthesis in water."),
    dict(doi="10.1038/s41586-020-2649-2", title="Array programming with NumPy",
         authors=_a("Harris|Charles R.", "Millman|K. Jarrod", "van der Walt|Stéfan J.", "Gommers|Ralf", "Virtanen|Pauli",
                    "Cournapeau|David", "Wieser|Eric", "Taylor|Julian", "Berg|Sebastian", "Smith|Nathaniel J.", "Kern|Robert",
                    "Picus|Matti", "Hoyer|Stephan", "van Kerkwijk|Marten H.", "Brett|Matthew", "Haldane|Allan",
                    "del Río|Jaime Fernández", "Wiebe|Mark", "Peterson|Pearu", "Gérard-Marchant|Pierre", "Sheppard|Kevin",
                    "Reddy|Tyler", "Weckesser|Warren", "Abbasi|Hameer", "Gohlke|Christoph", "Oliphant|Travis E."),
         year=2020, container="Nature", volume="585", issue="7825", pages="357-362", type="journal-article",
         issn=["0028-0836", "1476-4687"], publisher="Springer Science and Business Media LLC", language="en", cited_by=15000,
         oa_status="hybrid", source="crossref", abstract="NumPy array programming library for Python."),
    dict(doi="", title="自測示範條目：蛇紋石與深層海水文獻整理", authors=_a("自測示範|"), year=2026,
         container="自測示範刊物", volume="1", issue="1", pages="1-10", type="journal-article", language="zh",
         source="manual", note="使用者自填，未經 API 驗證（自測示範資料，非真實書目）"),
    dict(doi="", title="OpenAlex: The open catalog to the global research system", authors=_a("OurResearch|"), year=None,
         container="OpenAlex", type="webpage", url="https://openalex.org/", language="en", source="manual",
         note="使用者自填，未經 API 驗證（無卷期示範）"),
]
