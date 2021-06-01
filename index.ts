import { IInputs, IOutputs } from "./generated/ManifestTypes";

	import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
	type DataSet = ComponentFramework.PropertyTypes.DataSet;
	
	 //var optionsetMetaData: Array = [""];
	// Define const here
	const RowRecordId: string = "rowRecId";
	
	// Style name of Load More Button
	const LoadMoreButton_Hidden_Style = "LoadMoreButton_Hidden_Style";
	
	export class ColorGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	
		// Cached context object for the latest updateView
		private contextObj: ComponentFramework.Context<IInputs>;
	
		// Div element created as part of this control's main container
		private mainContainer: HTMLDivElement;
	
		// Table element created as part of this control's table
		private dataTable: HTMLTableElement;
	
		// Button element created as part of this control
		private loadPageButton: HTMLButtonElement;
	
	
		/**
		 * Empty constructor.
		 */
		constructor() {
	
		}

		public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
			// Need to track container resize so that control could get the available width. The available height won't be provided even this is true
			context.mode.trackContainerResize(true);
	//context.webAPI.retrieveMultipleRecords();
			// Create main table container div. 
			this.mainContainer = document.createElement("div");
			this.mainContainer.classList.add("SimpleTable_MainContainer_Style");
	
			// Create data table container div. 
			this.dataTable = document.createElement("table");
			this.dataTable.classList.add("SimpleTable_Table_Style");
	
			// Create data table container div. 
			this.loadPageButton = document.createElement("button");
			this.loadPageButton.setAttribute("type", "button");
			this.loadPageButton.innerText = context.resources.getString("PCF_ColorGrid_LoadMore_ButtonLabel");
			this.loadPageButton.classList.add(LoadMoreButton_Hidden_Style);
			this.loadPageButton.classList.add("LoadMoreButton_Style");
			this.loadPageButton.addEventListener("click", this.onLoadMoreButtonClick.bind(this));
	
			// Adding the main table and loadNextPage button created to the container DIV.
			this.mainContainer.appendChild(this.dataTable);
			this.mainContainer.appendChild(this.loadPageButton);
			container.appendChild(this.mainContainer);
		}
	
	
		/**
		 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
		 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
		 */
		public updateView(context: ComponentFramework.Context<IInputs>): void 
		{
			this.contextObj = context;
			this.toggleLoadMoreButtonWhenNeeded(context.parameters.simpleTableGrid);
	
			if (!context.parameters.simpleTableGrid.loading) {
	
				// Get sorted columns on View
				let columnsOnView = this.getSortedColumnsOnView(context);
	
				if (!columnsOnView || columnsOnView.length === 0) 
				{
					return;
				}
	
				let columnWidthDistribution = this.getColumnWidthDistribution(context, columnsOnView);
	
				var fieldValue = context.parameters.simpleTableGrid.addColumn? 'statuscode' : '';

				while (this.dataTable.firstChild) 
				{
					this.dataTable.removeChild(this.dataTable.firstChild);
				}
	
				this.dataTable.appendChild(this.createTableHeader(columnsOnView, columnWidthDistribution,context.parameters.headerBackgroundColor.raw!, context.parameters.headerForegroundColor.raw!));
				this.dataTable.appendChild(this.createTableBody(this, columnsOnView, columnWidthDistribution, context.parameters.simpleTableGrid, context.parameters.cellBackgroundColor.raw!, context.parameters.rowBackColor.raw!));
	
	
				this.dataTable.parentElement!.style.height = window.innerHeight - this.dataTable.offsetTop - 70 + "px";
			}
		}
	
		/** 
		 * It is called by the framework prior to a control receiving new data. 
		 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
		 */
		public getOutputs(): IOutputs {
			return {};
		}
	
		/** 
			 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
		 * i.e. cancelling any pending remote calls, removing listeners, etc.
		 */
		public destroy(): void {
		}
	
		/**
		 * Get sorted columns on view
		 * @param context 
		 * @return sorted columns object on View
		 */
		private getSortedColumnsOnView(context: ComponentFramework.Context<IInputs>): DataSetInterfaces.Column[] {
			if (!context.parameters.simpleTableGrid.columns) {
				return [];
			}
	
			let columns = context.parameters.simpleTableGrid.columns
				.filter(function (columnItem: DataSetInterfaces.Column) {
					// some column are supplementary and their order is not > 0
					return columnItem.order >= 0
				}
				);
	
			// Sort those columns so that they will be rendered in order
			columns.sort(function (a: DataSetInterfaces.Column, b: DataSetInterfaces.Column) {
				return a.order - b.order;
			});
	
			return columns;
		}
	
		/**
		 * Get column width distribution
		 * @param context context object of this cycle
		 * @param columnsOnView columns array on the configured view
		 * @returns column width distribution
		 */
		private getColumnWidthDistribution(context: ComponentFramework.Context<IInputs>, columnsOnView: DataSetInterfaces.Column[]): string[] {
	
			let widthDistribution: string[] = [];
	
			// Considering need to remove border & padding length
			let totalWidth: number = context.mode.allocatedWidth - 250;
			let widthSum = 0;
	
			columnsOnView.forEach(function (columnItem) {
				widthSum += columnItem.visualSizeFactor;
			});
	
			let remainWidth: number = totalWidth;
	
			columnsOnView.forEach(function (item, index) {
				let widthPerCell = "";
				if (index !== columnsOnView.length - 1) {
					let cellWidth = Math.round((item.visualSizeFactor / widthSum) * totalWidth);
					remainWidth = remainWidth - cellWidth;
					widthPerCell = cellWidth + "px";
				}
				else {
					widthPerCell = remainWidth + "px";
				}
				widthDistribution.push(widthPerCell);
			});
	
			return widthDistribution;
	
		}
	
	
		private createTableHeader(columnsOnView: DataSetInterfaces.Column[], widthDistribution: string[], headerBackgroundColor: string, headerForegroundColor: string): HTMLTableSectionElement {
	
			let tableHeader: HTMLTableSectionElement = document.createElement("thead");
			let tableHeaderRow: HTMLTableRowElement = document.createElement("tr");
			tableHeaderRow.classList.add("SimpleTable_TableRow_Style");
	
			tableHeaderRow.style.backgroundColor = headerBackgroundColor;
			tableHeaderRow.style.color = headerForegroundColor;
	
			columnsOnView.forEach(function (columnItem, index) {
				let tableHeaderCell = document.createElement("th");
				tableHeaderCell.classList.add("SimpleTable_TableHeader_Style");
				let innerDiv = document.createElement("div");
				innerDiv.classList.add("SimpleTable_TableCellInnerDiv_Style");
				innerDiv.style.maxWidth = widthDistribution[index];
				innerDiv.innerText = columnItem.displayName;
				tableHeaderCell.appendChild(innerDiv);
				tableHeaderRow.appendChild(tableHeaderCell);
			});
	
			tableHeader.appendChild(tableHeaderRow);
			return tableHeader;
		}
	
		private createTableBody(thisControl: any, columnsOnView: DataSetInterfaces.Column[], widthDistribution: string[], gridParam: DataSet, cellBackgroundColorExpression: string, rowBackgroundColor: string): HTMLTableSectionElement {
	
			let tableBody: HTMLTableSectionElement = document.createElement("tbody");
			let rowNumber: number = 0;
	
			if (gridParam.sortedRecordIds.length > 0) 
			{
				for (let currentRecordId of gridParam.sortedRecordIds) 
				{
					let tableRecordRow: HTMLTableRowElement = document.createElement("tr");
					tableRecordRow.classList.add("SimpleTable_TableRow_Style");
					let rowBackgroundColor1 = '';
					try 
					{
						rowBackgroundColor1 = "=((ROWNUMBER() % 2) == fieldSchema) ? 'lightblue' : 'white';"
						/*if (rowBackgroundColor.startsWith("=")) 
						{
							let rowBackColorExpression = rowBackgroundColor.substring(1);
							rowBackColorExpression = rowBackColorExpression.replace('ROWNUMBER()', 'rowNumber');
							rowBackgroundColor1 = eval(rowBackColorExpression);
						}
						else 
						{
							rowBackgroundColor1 = rowBackgroundColor;
						}
						console.log(rowBackgroundColor1);*/
						//alert("statusValue " + statusValue);  
					
						let statusValue:number =+gridParam.records[currentRecordId].getValue('statuscode');
						
									switch(statusValue)
									{
										case 1: 
										{
											rowBackgroundColor1 = 'green';
										}
										case 2:
										{
											rowBackgroundColor1 = 'red';
										}
									}
						tableRecordRow.style.backgroundColor = rowBackgroundColor1;
					}
					catch (err) 
					{
						console.log("Invalid Row Background Color expression");
						console.log(err.message);
					}
					rowNumber++;
	
	
					//tableRecordRow.addEventListener("click", this.onRowClick.bind(this));
	
					// Set the recordId on the row dom
					tableRecordRow.setAttribute(RowRecordId, gridParam.records[currentRecordId].getRecordId());
	
	
					columnsOnView.forEach(function (columnItem, index) {
						let tableRecordCell = document.createElement("td");
						tableRecordCell.classList.add("SimpleTable_TableCell_Style");
	
	
						//Single Color Expression for background color implementation. 
						var fieldcolorExpression = cellBackgroundColorExpression;
						var fieldname = '';
						var backgroundColor = 'green';
	/*
						//check whether it is an expression. 
						if (fieldcolorExpression.startsWith("=")) 
						{
							//Check whether the expression has fieldname
							if (fieldcolorExpression.indexOf(fieldname) > -2) 
							{
								//Remove =
								fieldcolorExpression = fieldcolorExpression.replace('=', '');
	
								//Get field name
								let start = fieldcolorExpression.indexOf("('") + 2;
								let finish = fieldcolorExpression.indexOf("')");
								let fieldname = fieldcolorExpression.substring(start, finish);
								console.log(fieldname);
	
								if (columnItem.name == fieldname) 
								{
									//Replace GETFIELDVALUE by dataset getValue
									//fieldcolorExpression = fieldcolorExpression.replace('GETFIELDVALUE', 'gridParam.records[currentRecordId].getValue');
									fieldcolorExpression = gridParam.records[currentRecordId].getValue('statuscode').toString();					
									console.log(fieldcolorExpression);									
	
									try 
									{
										if(eval(fieldcolorExpression) == 1)
										{
											backgroundColor = eval(fieldcolorExpression);
											console.log(backgroundColor);
											tableRecordCell.style.backgroundColor = 'green';
										}
										if(eval(fieldcolorExpression) == 2)
										{
											backgroundColor = eval(fieldcolorExpression);
											console.log(backgroundColor);
											tableRecordCell.style.backgroundColor = 'indianred';
										}											
									}
									catch (err) {
										console.log("Invalid field color expression" + fieldcolorExpression);
										console.log(err.message);
										return;
									}
								}
							}
						}
						else 
						{//No expression,user has set hardcoded color in property
							if (fieldcolorExpression != '') 
							{
								backgroundColor = fieldcolorExpression;
								console.log(backgroundColor);
								tableRecordCell.style.backgroundColor = backgroundColor;
							}
						}*/
						let statusValue:number =+gridParam.records[currentRecordId].getValue('statuscode');
						//var optionset 
						if(statusValue == 1)
						{

							tableRecordCell.style.backgroundColor = 'green';
						}
						else if(statusValue == 2)
						{
							tableRecordCell.style.backgroundColor = 'lightblue';
						}
						else
						{
							tableRecordCell.style.backgroundColor = 'white';
						}									
						let innerDiv = document.createElement("div");
						innerDiv.classList.add("SimpleTable_TableCellInnerDiv_Style");
						innerDiv.style.maxWidth = widthDistribution[index];
						innerDiv.innerText = gridParam.records[currentRecordId].getFormattedValue(columnItem.name);
						tableRecordCell.appendChild(innerDiv);
						tableRecordRow.appendChild(tableRecordCell);
					});
	
					tableBody.appendChild(tableRecordRow);
				}
			}
			else {
				let tableRecordRow: HTMLTableRowElement = document.createElement("tr");
				let tableRecordCell: HTMLTableCellElement = document.createElement("td");
				tableRecordCell.classList.add("No_Record_Style");
				tableRecordCell.colSpan = columnsOnView.length;
				tableRecordCell.innerText = this.contextObj.resources.getString("PCF_ColorGrid_No_Record_Found");
				tableRecordRow.appendChild(tableRecordCell)
				tableBody.appendChild(tableRecordRow);
			}
	
			return tableBody;
		}
	
	 
		/**
		 * Row Click Event handler for the associated row when being clicked
		 * @param event
		 
		private onRowClick(event: Event): void 
		{
			let rowRecordId = (event.currentTarget as HTMLTableRowElement).getAttribute(RowRecordId);
	
			if (rowRecordId) 
			{
				let entityreference = this.contextObj.parameters.simpleTableGrid.records[rowRecordId].getNamedReference();
				  let entityFormOptions = 
				  {
					entityName: entityreference.name,
					entityId: entityreference.id
				  };
				  this.contextObj.navigation.openForm(entityFormOptions);
			}
		}*/
	
		/**
		 * Toggle 'LoadMore' button when needed
		 */
		private toggleLoadMoreButtonWhenNeeded(gridParam: DataSet): void {
	
			if (gridParam.paging.hasNextPage && this.loadPageButton.classList.contains(LoadMoreButton_Hidden_Style)) {
				this.loadPageButton.classList.remove(LoadMoreButton_Hidden_Style);
			}
			else if (!gridParam.paging.hasNextPage && !this.loadPageButton.classList.contains(LoadMoreButton_Hidden_Style)) {
				this.loadPageButton.classList.add(LoadMoreButton_Hidden_Style);
			}
	
		}
	
		/**
		 * 'LoadMore' Button Event handler when load more button clicks
		 * @param event
		 */
		private onLoadMoreButtonClick(event: Event): void {
			this.contextObj.parameters.simpleTableGrid.paging.loadNextPage();
			this.toggleLoadMoreButtonWhenNeeded(this.contextObj.parameters.simpleTableGrid);
		}
	
	}