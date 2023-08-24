<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<!-- xmlns:xsl="www.w3.org/1999/XSL/Transform" -->
	<xsl:template match="/">
		<CreateContractQuests>
			<TestCase>
				<DatasetName><xsl:value-of select="//DatasetName"/></DatasetName>
				<GeneratedAt><xsl:value-of select="//GeneratedAt"/></GeneratedAt>
			</TestCase>
			<StudyResults>
				<xsl:for-each select="dataset/data/StudyResult">
					<StudyResult>
						<StudyQuests>
							<xsl:for-each select="row">
								<StudyQuest>
									<QuestId><xsl:value-of select="@id"/></QuestId>
								</StudyQuest>
							</xsl:for-each>
						</StudyQuests>
					</StudyResult>
				</xsl:for-each>
			</StudyResults>
		</CreateContractQuests>
	</xsl:template>
</xsl:stylesheet>
