<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="2.0">
<xsl:template match="/">
<![CDATA[
<?xml version="1.0" encoding="UTF-8"?>
<pma_xml_export version="1.0" xmlns:pma="http://www.phpmyadmin.net/some_doc_url/">
    <!--
    - Structure schemas
    -->
    <pma:structure_schemas>
        <pma:database name="sozluk" collation="latin1_swedish_ci" charset="latin1">
            <pma:table name="dictionary">
                CREATE TABLE `dictionary` (
                  `id` int(11) NOT NULL AUTO_INCREMENT,
                  `latin` varchar(255) COLLATE utf8_unicode_ci NOT NULL,
                  `kiril` varchar(255) COLLATE utf8_unicode_ci NOT NULL,
                  `mean` longtext COLLATE utf8_unicode_ci NOT NULL,
                  PRIMARY KEY (`id`)
                ) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
            </pma:table>
        </pma:database>
    </pma:structure_schemas>
	
    <!--
    - Database: 'sozluk'
    -->
    <database name="sozluk">
]]>

<xsl:for-each select="stardict/article">
<![CDATA[<table name="dictionary">]]><br/>
		<![CDATA[<column name="id">]]>NULL<![CDATA[</column>]]><br/>
		<![CDATA[<column name="latin">]]><xsl:value-of select="key"/><![CDATA[</column>]]><br/>
		<![CDATA[<column name="kiril">]]>NULL<![CDATA[</column>]]><br/>
		<![CDATA[<column name="mean">[CDATA[]]>
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