<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="2.0">
<xsl:template match="/">
<![CDATA[
<?xml version="1.0" encoding="UTF-8"?>
<pma_xml_export version="1.0" xmlns:pma="http://www.phpmyadmin.net/some_doc_url/">
	<pma:structure_schemas>
        <pma:database name="dictionaries" collation="utf8_general_ci" charset="utf8">
            <pma:table name="dict_cumakunova_35_200">
                CREATE TABLE `dict_cumakunova_35_200` (
                  `id` int(11) NOT NULL AUTO_INCREMENT,
                  `term` varchar(1500) NOT NULL,
                  `definition` text NOT NULL,
                  PRIMARY KEY (`id`)
                ) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8;
            </pma:table>
        </pma:database>
    </pma:structure_schemas>	
<database name="baktybek2_ktmu">
]]>

<xsl:for-each select="stardict/article">
<![CDATA[<table name="dict_cumakunova_35_200">]]><br/>
		<![CDATA[<column name="id">]]>NULL<![CDATA[</column>]]><br/>
		<![CDATA[<column name="term">]]><xsl:value-of select="key"/><![CDATA[</column>]]><br/>
		<![CDATA[<column name="definition">[CDATA[]]>
		<xsl:value-of select="definition"/><br/>
		<![CDATA[]]</column>]]><br/>
<![CDATA[</table>]]><br/>
</xsl:for-each>
<![CDATA[
</database>
</pma_xml_export>
]]>
</xsl:template>
</xsl:stylesheet>