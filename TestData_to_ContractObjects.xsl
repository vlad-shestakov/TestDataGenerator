<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
				xmlns:i="http://www.w3.org/2001/XMLSchema-instance"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:template match="/">
	<GetContractObjectsByIDRequestsResponse>
		<GetContractObjectsByIDResult xmlns:i="http://www.w3.org/2001/XMLSchema-instance">
			<Objects>


				<!-- Строки с NIL_VALUE ("nil") будут преобразованы в тег </tag i:nil="true"> -->
				<xsl:variable name="NIL_VALUE" select="'nil'" />

				<!-- =========== Перебор тест-кейсов (Вкладок Case1 / Case2) =========== -->
				<!-- !!!!! Только первая вкладка - Case1 !!!!!!!  -->
				<xsl:for-each select="dataset/data[1]">

					<ContractObjects>


					<!-- =============== ContractObject/row ===============  -->
					<xsl:for-each select="ContractObject/row">

						<!-- Запомним в ContractObjectRowId идентификатор строки (Qest1_1, Qest1_2).
							 Для таблицы ContractObject это 'Qest1' + порядковый номер строки _N -->
						<xsl:variable name="ContractObjectRowId" select="@id" />

						<ContractObject>
							<GeneralBlock>

								<!-- <Number>35001007</Number> -->
								<RequestNumber><xsl:value-of select="../../Contract/row/Field[@name='RequestNumber']"/></RequestNumber>

								<!-- <ActualDocNumber>12 34 123456</ActualDocNumber> -->
								<ActualDocNumber><xsl:value-of select="Field[@name='ActualDocNumber']"/></ActualDocNumber>
								<!-- <BirthDate>01.01.1980</BirthDate> -->
								<BirthDate><xsl:value-of select="Field[@name='BirthDate']"/></BirthDate>
								<!-- <ObjectName>ЗаемщикОдин Николя Петрович</ObjectName> -->
								<ObjectName><xsl:value-of select="Field[@name='ObjectName']"/></ObjectName>
								<!-- <GenderCode>Male</GenderCode> -->
								<GenderCode><xsl:value-of select="Field[@name='GenderCode']"/></GenderCode>
								<!-- <InsuranceLifePercent>50</InsuranceLifePercent> -->
								<InsPercent><xsl:value-of select="Field[@name='InsPercent']"/></InsPercent>
								<!-- <RoleType>Borrower</RoleType> -->
								<RoleType><xsl:value-of select="Field[@name='Role']"/></RoleType>
							</GeneralBlock>


							<InfoBlock>

								<!-- =============== InfoBlock/row ===============  -->
								<!-- Выберем строки по полю InfoBlock.REF_MasterRecord = $ContractObjectRowId -->
								<xsl:for-each select="../../InfoBlock/row[./Field[@name='REF_MasterRecord']=$ContractObjectRowId]">

									<!-- <IsSmokingNow>false</IsSmokingNow> -->
									<xsl:choose><xsl:when test="./Field[@name='IsSmokingNow']/text()=$NIL_VALUE"><IsSmokingNow i:nil="true" /></xsl:when>
										<xsl:otherwise><IsSmokingNow><xsl:value-of select="Field[@name='IsSmokingNow']"/></IsSmokingNow></xsl:otherwise></xsl:choose>
									<!-- <SmokingPeriod i:nil="true"/> -->
									<xsl:choose><xsl:when test="./Field[@name='SmokingPeriod']/text()=$NIL_VALUE"><SmokingPeriod i:nil="true" /></xsl:when>
										<xsl:otherwise><SmokingPeriod><xsl:value-of select="Field[@name='SmokingPeriod']"/></SmokingPeriod></xsl:otherwise></xsl:choose>

								</xsl:for-each>
								<!-- /=============== InfoBlock/row ===============  -->
							</InfoBlock>

							<WorkplaceBlock>
								<!-- <JobTitle>Продавец-кассир</JobTitle> -->
								<JobTitle><xsl:value-of select="Field[@name='JobTitle']"/></JobTitle>
							</WorkplaceBlock>
						</ContractObject>

					</xsl:for-each>
					<!-- /=============== ContractObject/row ===============  -->

					</ContractObjects>

				</xsl:for-each>
				<!-- / =========== Перебор тест-кейсов (Вкладок Case1 / Case2) =========== -->

			</Objects>
		</GetContractObjectsByIDResult>
	</GetContractObjectsByIDRequestsResponse>


</xsl:template>
</xsl:stylesheet>
