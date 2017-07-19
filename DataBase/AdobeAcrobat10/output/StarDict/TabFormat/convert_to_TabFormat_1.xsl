<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="2.0">
<xsl:template match="/">

<xsl:for-each select="stardict/article">

		<![CDATA[entry_start]]>	
		<xsl:value-of select="key"/>
		<![CDATA[Tab_is_here]]>	
		<xsl:value-of select="definition"/>
		<![CDATA[entry_end]]>			
		
</xsl:for-each>
</xsl:template>
</xsl:stylesheet>
