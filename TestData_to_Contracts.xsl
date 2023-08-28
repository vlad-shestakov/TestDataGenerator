<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:template match="/">
				
<GetAllContractsRequestsResponse>
	<GetAllContractsRequestsResult xmlns:i="http://www.w3.org/2001/XMLSchema-instance">
		<Contracts>

		<!-- =========== Перебор тест-кейсов (Вкладок Case1 / Case2) =========== -->
		<xsl:for-each select="dataset/data">

            <Contract>
                <!-- <Number>35001007</Number> -->
                <Number><xsl:value-of select="Contract/row/Field[@name='RequestNumber']"/></Number>
                <!-- <CreditId>a31ede1b-ef6e-ed11-84d3-005056bf5af5</CreditId> -->
                <CreditId><xsl:value-of select="Contract/row/Field[@name='CreditId']"/></CreditId>
                <!-- <CreditAmount>2000000.00</CreditAmount> -->
                <CreditAmount><xsl:value-of select="Contract/row/Field[@name='CreditAmount']"/></CreditAmount>
                <!-- <CreditPercent>10.4500000000</CreditPercent> -->
                <CreditPercent><xsl:value-of select="Contract/row/Field[@name='CreditPercent']"/></CreditPercent>
                <!-- <CreditTerm>60</CreditTerm> -->
                <CreditTerm><xsl:value-of select="Contract/row/Field[@name='CreditTerm']"/></CreditTerm>

					<ContractObjects>
				  
					<!-- =============== ContractObject/row ===============  -->
					<xsl:for-each select="ContractObject/row">
						
                     <ContractObject>
                         <!-- <ObjectId>cdb3a837-ef6e-ed11-84d3-005056bf5ab80</ObjectId> -->
                         <ObjectId><xsl:value-of select="Field[@name='ObjectId']"/></ObjectId>
                         <!-- <InsPercent>50</InsPercent> -->
                         <InsPercent><xsl:value-of select="Field[@name='InsPercent']"/></InsPercent>
                        <!-- <Name>ЗаемщикОдин Николя Петрович</Name> -->
                        <Name><xsl:value-of select="Field[@name='ObjectName']"/></Name>
                        <!-- <Role>BORROWER</Role> -->
                        <Role><xsl:value-of select="Field[@name='Role']"/></Role>
                     </ContractObject>
					 
					</xsl:for-each>
					<!-- /=============== ContractObject/row ===============  -->
					
                    </ContractObjects>
                    <!--State>NEW</State-->
                    <State><xsl:value-of select="Contract/row/Field[@name='State']"/></State>
			</Contract>

		</xsl:for-each>
		<!-- / =========== Перебор тест-кейсов (Вкладок Case1 / Case2) =========== -->
		
		</Contracts>
	</GetAllContractsRequestsResult>
</GetAllContractsRequestsResponse>
</xsl:template>
</xsl:stylesheet>