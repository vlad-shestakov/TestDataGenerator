<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<!-- xmlns:xsl="www.w3.org/1999/XSL/Transform" -->
	<xsl:template match="/">
		<my_data>
			<xsl:for-each select="data/e">
				<xsl:value-of select="@v"/>
			</xsl:for-each>
		</my_data>
	</xsl:template>
</xsl:stylesheet>
