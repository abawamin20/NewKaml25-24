import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import "@pnp/sp/search";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { CellRender } from "../common/ColumnDetails";
import { getColumnMaxWidth, getColumnMinWidth } from "../utils/columnUtils";
import { ConstructedFilter } from "./PanelComponent";
export interface ITerm {
  Id: string;
  Name: string;
  parentId: string;
  Children?: ITerm[];
}

export interface TermSet {
  setId: string;
  terms: ITerm[];
}

export interface FilterDetail {
  filterColumn: string;
  filterColumnType: string;
  values: string[];
}

export interface IColumnInfo {
  InternalName: string;
  DisplayName: string;
  MinWidth: number;
  ColumnType: string;
  MaxWidth: number;
  OnRender?: (items: any) => JSX.Element;
}
class PagesService {
  private _sp: SPFI;

  constructor(private context: WebPartContext) {
    this._sp = spfi().using(SPFx(this.context));
  }

  /**
   * Fetch distinct values for a given column from a list of items.
   * @param {string} columnName - The name of the column to fetch distinct values for.
   * @param {any[]} values - The list of items to extract distinct values from.
   * @returns {Promise<string[] | ConstructedFilter[]>} - A promise that resolves to an array of distinct values.
   */
  getDistinctValues = async (
    columnName: string,
    columnType: string,
    values: any
  ): Promise<(string | ConstructedFilter)[]> => {
    try {
      const items = values; // The list of items to fetch distinct values from.

      // Extract distinct values from the column
      const distinctValues: (string | ConstructedFilter)[] = [];
      const seenValues = new Set<string | ConstructedFilter>(); // A set to keep track of seen values to avoid duplicates.

      items.forEach((item: any) => {
        switch (columnType) {
          case "TaxonomyFieldTypeMulti":
            if (item[columnName] && item[columnName].length > 0) {
              // Extract distinct values from the column
              item[columnName].forEach((category: any) => {
                const uniqueValue = category.Label;
                if (!seenValues.has(uniqueValue)) {
                  seenValues.add(uniqueValue);
                  distinctValues.push(uniqueValue);
                }
              });
            }
            break;
          case "DateTime":
            let uniqueDateValue = item[columnName]; // The value of the column for the current item.
            // Handle ISO date strings by extracting only the date part
            uniqueDateValue = new Date(uniqueDateValue)
              .toISOString()
              .split("T")[0];

            if (!seenValues.has(uniqueDateValue)) {
              seenValues.add(uniqueDateValue);
              distinctValues.push(uniqueDateValue);
            }
            break;
          case "User":
            const userValue = item[columnName];
            if (
              userValue &&
              userValue.Title &&
              !seenValues.has(userValue.Title)
            ) {
              seenValues.add(userValue.Title);
              const user: ConstructedFilter = {
                text: userValue.Title,
                value: userValue.Id,
              };
              distinctValues.push(user);
            }
            break;
          case "Number":
            const uniqueNumberValue = item[columnName]; // The value of the column for the current item.

            if (!seenValues.has(uniqueNumberValue)) {
              seenValues.add(uniqueNumberValue);
              distinctValues.push(uniqueNumberValue);
            }
            break;
          case "Choice":
            const uniqueChoiceValue = item[columnName]; // The value of the column for the current item.
            if (uniqueChoiceValue) {
              if (!seenValues.has(uniqueChoiceValue)) {
                seenValues.add(uniqueChoiceValue);
                distinctValues.push(uniqueChoiceValue);
              }
            }
            break;
          case "URL":
            const uniqueUrlChoiceValue = item[columnName]; // The value of the column for the current item.
            if (uniqueUrlChoiceValue && uniqueUrlChoiceValue.Url) {
              if (!seenValues.has(uniqueUrlChoiceValue.Url)) {
                seenValues.add(uniqueUrlChoiceValue.Url);
                distinctValues.push(uniqueUrlChoiceValue.Url);
              }
            }
            break;
          case "Computed":
            const uniqueCompChoiceValue = item[columnName]; // The value of the column for the current item.
            if (uniqueCompChoiceValue) {
              if (!seenValues.has(uniqueCompChoiceValue.split(".")[0])) {
                seenValues.add(uniqueCompChoiceValue.split(".")[0]);
                distinctValues.push(uniqueCompChoiceValue.split(".")[0]);
              }
            }
            break;
          default:
            const uniqueValue = item[columnName]; // The value of the column for the current item.
            if (uniqueValue) {
              if (!seenValues.has(uniqueValue)) {
                seenValues.add(uniqueValue);
                distinctValues.push(uniqueValue);
              }
            }
            break;
        }
      });

      return distinctValues;
    } catch (error) {
      console.error("Error fetching distinct values:", error);
      throw error;
    }
  };

  getDistinctValues2 = async (
    listId: string,
    columnName: string
  ): Promise<any[]> => {
    try {
      // Build the RenderListFilterData endpoint URL
      const filterDataEndpoint = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/RenderListFilterData.aspx?FieldInternalName=${columnName}&ListId=${listId}`;

      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          filterDataEndpoint,
          SPHttpClient.configurations.v1
        );
      const responseData = await response.json();

      // Extract distinct values
      const distinctValues = responseData.filterData.map((item: any) => ({
        text: item.Value, // Display value
        value: item.Key, // Internal value
      }));

      return distinctValues;
    } catch (error) {
      console.error("Error fetching distinct values:", error);
      throw error;
    }
  };

  /**
   * Retrieves a page of filtered Site Pages items.
   *
   * @param orderBy The column to sort the items by. Defaults to "Created".
   * @param isAscending Whether to sort in ascending or descending order. Defaults to true.
   * @param category The current category that is selected by the user
   * @param searchText Text to search for in the Title, Article ID, or Modified columns.
   * @param filters An array of FilterDetail objects to apply to the query.
   * @param columnInfos The columns information to include in the query.
   * @param pagesCache Cache to store fetched pages.
   * @param currentPageIndex The current page index.
   * @returns A promise that resolves with an array of items.
   */
  getFilteredPages = async (
    orderBy: string = "Created",
    isAscending: boolean = true,
    category: string = "",
    searchText: string = "",
    filters: FilterDetail[],
    columnInfos: IColumnInfo[],
    pageSize: number,
    pageIndex: number = 0 // Added pageIndex parameter
  ) => {
    try {
      console.log(columnInfos);
      const expandFieldsSet = new Set<string>();

      {
        // Default columns to always include
        expandFieldsSet.add("FieldValuesAsText");

        // Build CAML Query
        let filterConditions = `
<And>
  <Eq>
    <FieldRef Name='KnowledgeBaseLabel' />
    <Value Type='Text'>${category}</Value>
  </Eq>
  <Eq>
    <FieldRef Name='FSObjType' />
    <Value Type='Integer'>0</Value>
  </Eq>
</And>`;

        const additionalFilters = [];

        // Add search text filtering
        if (searchText) {
          additionalFilters.push(`
  <Or>
    <Contains>
      <FieldRef Name='Title' />
      <Value Type='Text'>${searchText}</Value>
    </Contains>
    <Eq>
      <FieldRef Name='Article_x0020_ID' />
      <Value Type='Text'>${searchText}</Value>
    </Eq>
    <Contains>
      <FieldRef Name='Modified' />
      <Value Type='DateTime'>${searchText}</Value>
    </Contains>
  </Or>`);
        }

        // Apply additional filters from the filter array
        filters.forEach((filter) => {
          if (filter.values.length > 0) {
            let filterClauses: string[] = [];
            switch (filter.filterColumnType) {
              case "DateTime":
                filter.values.forEach((value) => {
                  const startDate = new Date(value).toISOString();
                  const endDate = new Date(value);
                  endDate.setDate(endDate.getDate() + 1);
                  const endDateStr = endDate.toISOString();

                  filterClauses.push(`
            <And>
              <Geq>
                <FieldRef Name='${filter.filterColumn}' />
                <Value Type='DateTime'>${startDate}</Value>
              </Geq>
              <Lt>
                <FieldRef Name='${filter.filterColumn}' />
                <Value Type='DateTime'>${endDateStr}</Value>
              </Lt>
            </And>`);
                });
                break;

              case "User":
                filter.values.forEach((value) => {
                  filterClauses.push(`
            <Eq>
              <FieldRef Name='${filter.filterColumn}' LookupId='TRUE' />
              <Value Type='User'>${value}</Value>
            </Eq>`);
                });
                break;

              case "URL":
                filter.values.forEach((value) => {
                  filterClauses.push(`
            <Eq>
              <FieldRef Name='${filter.filterColumn}' />
              <Value Type='URL'>${value}</Value>
            </Eq>`);
                });
                break;

              default:
                filter.values.forEach((value) => {
                  filterClauses.push(`
            <Eq>
              <FieldRef Name='${filter.filterColumn}' />
              <Value Type='Text'>${value}</Value>
            </Eq>`);
                });
                break;
            }

            // If there is more than one value, wrap in <Or> tags
            if (filterClauses.length > 1) {
              additionalFilters.push(`<Or>${filterClauses.join("")}</Or>`);
            } else {
              additionalFilters.push(filterClauses[0]);
            }
          }
        });

        // Combine filterConditions with additionalFilters
        if (additionalFilters.length > 0) {
          filterConditions = `<And>${filterConditions}${additionalFilters.join(
            ""
          )}</And>`;
        }

        // Order by clause
        const orderByClause = `
        <OrderBy>
          <FieldRef Name='${orderBy}' Ascending='${isAscending}' />
        </OrderBy>`;

        // Construct CAML Query
        const camlQuery = `
        <View Scope='RecursiveAll'>
          <Query>
            <Where>${filterConditions}</Where>
            ${orderByClause}
          </Query>
          <RowLimit>${pageSize}</RowLimit>
        </View>`;
        // Use pageIndex and pagesSize to calculate the number of items to skip
        const skip = pageIndex * pageSize;

        const body = {
          query: {
            __metadata: {
              type: "SP.CamlQuery",
            },
            ViewXml: camlQuery,
            ListItemCollectionPosition: {
              PagingInfo: `${skip}`, // Skip to the correct page
            },
          },
        };

        const expandedFields: string[] = [];

        expandFieldsSet.forEach((field) => expandedFields.push(field));
        // Fetch items using CAML Query
        const response: SPHttpClientResponse =
          await this.context.spHttpClient.post(
            `${
              this.context.pageContext.web.absoluteUrl
            }/_api/web/lists/getByTitle('Site Pages')/GetItems?$expand=${expandedFields.join()}&$top=${pageSize}&$skiptoken=${skip}`,
            SPHttpClient.configurations.v1,
            {
              body: JSON.stringify(body),
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "odata-version": "",
              },
            }
          );
        const jsonResponse = await response.json();

        // Next page URL handling
        // Check if there are more items to paginate
        const moreItems = jsonResponse.d.results.length >= pageSize;
        return {
          pages: jsonResponse.d.results,
          nextPageIndex: moreItems ? pageIndex + 1 : null, // Increment pageIndex for next page
        };
      }
    } catch (error) {
      console.error(error);
    }
  };

  /**
   * Retrieves the columns for a specified view in the SharePoint list.
   */
  public async getColumns(viewId: string): Promise<IColumnInfo[]> {
    const fields = await this._sp.web.lists
      .getByTitle("Site Pages")
      .views.getById(viewId)
      .fields();

    // Fetching detailed field information to get both internal and display names
    const fieldDetailsPromises = fields.Items.map((field: any) =>
      this._sp.web.lists
        .getByTitle("Site Pages")
        .fields.getByInternalNameOrTitle(field)()
    );

    const fieldDetails = await Promise.all(fieldDetailsPromises);

    return fieldDetails.map((field: any) => ({
      InternalName: field.InternalName,
      DisplayName: field.Title,
      ColumnType: field.TypeAsString,
      MinWidth: getColumnMinWidth(field.InternalName),
      MaxWidth: getColumnMaxWidth(field.InternalName),
      OnRender: (item: any) =>
        CellRender({
          columnName: field.InternalName,
          columnType: field.TypeAsString,
          item,
          context: this.context,
        }),
    }));
  }

  /**
   * Retrieves the details of a SharePoint list by its name.
   * @param {string} listName - The name of the list to retrieve details for.
   * @returns {Promise<any>} - A promise that resolves to the list details.
   */
  public async getListDetailsByName(listName: string): Promise<any> {
    try {
      const list = await this._sp.web.lists.getByTitle(listName)();
      return list;
    } catch (error) {
      console.error(`Error retrieving list details for ${listName}:`, error);
      throw new Error(`Error retrieving list details for ${listName}`);
    }
  }

  public async createListItem(itemData: any, listTitle: string): Promise<any> {
    try {
      const addedItem = await this._sp.web.lists
        .getByTitle(listTitle) // Get list by title
        .items.add(itemData);
      return addedItem;
    } catch (error) {
      console.error("Error creating list item: ", error);
      throw error;
    }
  }

  async getByUrl(url: string): Promise<any> {
    try {
      return this.context.spHttpClient
        .get(url, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        });
    } catch (error) {}
  }
}

export default PagesService;
