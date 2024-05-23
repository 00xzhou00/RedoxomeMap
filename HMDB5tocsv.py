import pandas as pd
from lxml import etree
def hmdbtocsv(xmlfile,csvfile):
    accessions=[]
    monisotopic_molecular_weights=[]
    iupac_names=[]
    names=[]
    chemical_formulas=[]
    keggs=[]
    ns = {'hmdb': 'http://www.hmdb.ca'}
    context = etree.iterparse(xmlfile, tag='{http://www.hmdb.ca}metabolite')
    for event, elem in context:

        accession = elem.xpath('hmdb:accession/text()', namespaces=ns)[0]
        try:
            monisotopic_molecular_weight = elem.xpath('hmdb:monisotopic_molecular_weight/text()', namespaces=ns)[0]
        except:
            monisotopic_molecular_weight = 'NA'
        try:
            iupac_name = elem.xpath('hmdb:iupac_name/text()', namespaces=ns)[0].encode('utf-8')
        except:
            iupac_name = 'NA'
        name = elem.xpath('hmdb:name/text()', namespaces=ns)[0].encode('utf-8')
        try:
            chemical_formula = elem.xpath('hmdb:chemical_formula/text()', namespaces=ns)[0]
        except:
            chemical_formula = 'NA'
        try:
            kegg = elem.xpath('hmdb:kegg_id/text()', namespaces=ns)[0]
        except:
            kegg = 'NA'
        accessions.append(accession)
        monisotopic_molecular_weights.append(monisotopic_molecular_weight)
        iupac_names.append(iupac_name)
        names.append(name)
        chemical_formulas.append(chemical_formula)
        keggs.append(kegg)
        elem.clear()

        for ancestor in elem.xpath('ancestor-or-self::*'):

            while ancestor.getprevious() is not None:

                del ancestor.getparent()[0]
    del context
    outdf = pd.DataFrame({'accession':accessions,'monisotopic_molecular_weight':monisotopic_molecular_weights,'iupac_name':iupac_names,'name':names,'chemical_formula':chemical_formulas,'kegg':keggs})
    outdf.to_csv(csvfile, index=False)