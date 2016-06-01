<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="2.0">
<xsl:template match="/">

<xsl:for-each select="stardict/article">
<br/>
<br/>
		<xsl:value-of select="key"/>
		<![CDATA[Hello]]>	
		<xsl:value-of select="definition"/>
		

</xsl:for-each>
</xsl:template>
</xsl:stylesheet>