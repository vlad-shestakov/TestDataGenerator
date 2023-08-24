<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<!-- xmlns:xsl="www.w3.org/1999/XSL/Transform" -->
	<xsl:template match="/">
		<CreateContract>
			<TestCase>
				<DatasetName><xsl:value-of select="//DatasetName"/></DatasetName>
				<GeneratedAt><xsl:value-of select="//GeneratedAt"/></GeneratedAt>
			</TestCase>
			<StudyResults>
				<xsl:for-each select="dataset/data/StudyResult">
					<StudyResult>
						<ResultType><xsl:value-of select="@type"/></ResultType>
					</StudyResult>
				</xsl:for-each>
			</StudyResults>
		</CreateContract>
	</xsl:template>
</xsl:stylesheet>