import * as React from "react";
import { Dialog, DialogFooter, Spinner } from "@fluentui/react";
import { SPHttpClient } from "@microsoft/sp-http";
import { ReusableDetailList } from "../common/ReusableDetailList";
import PagesService, { FilterDetail, IColumnInfo } from "./PagesService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PagesColumns } from "./PagesColumns";
import { DefaultButton, IColumn, Icon, Selection } from "@fluentui/react";
import { makeStyles, useId, Input } from "@fluentui/react-components";
import styles from "./pages.module.scss";
import "./pages.css";
import { FilterPanelComponent } from "./PanelComponent";
import ListForm from "../Forms/ListForm";

export interface IPagesListProps {
  context: WebPartContext;
  selectedViewId: string;
  feedbackPageUrl: string;
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    gap: "2px",
    maxWidth: "400px",
    alignItems: "center",
  },
});

const PagesList = (props: IPagesListProps) => {
  const subscribeLink: string = "/_layouts/15/SubNew.aspx";
  const alertLink: string = "/_layouts/15/mySubs.aspx";

  // Destructure the props
  const { context, selectedViewId } = props;

  /**
   * State variables for the component.
   */

  // Options for the page size dropdown
  const [pageSizeOption] = React.useState<number[]>([
    10, 20, 40, 60, 80, 100, 200, 300, 400, 500,
  ]);

  const [hideFeedBackDialog, setHideFeedBackDialog] = React.useState(true);

  const toggleHideFeedbackDialog = () => {
    setHideFeedBackDialog(!hideFeedBackDialog);
  };
  const [hideAlertMeDialog, setHideAlertMeDialog] = React.useState(true);
  const [hideManageAlertDialog, setHideManageAlertDialog] =
    React.useState(true);

  const toggleHideAlertMeDialog = () => {
    setHideAlertMeDialog(!hideAlertMeDialog);
  };
  const toggleHideManageAlertDialog = () => {
    setHideManageAlertDialog(!hideManageAlertDialog);
  };

  const [columnInfos, setColumnInfos] = React.useState<IColumnInfo[]>([]);

  // The search text for filtering pages
  const [searchText, setSearchText] = React.useState<string>(""); // Initially set to empty string

  // The list of pages
  const [pages, setPages] = React.useState<any[]>([]); // Initially set to empty array

  // The selected category
  const [catagory, setCatagory] = React.useState<string | null>(null); // Initially set to empty string
  const [isLoading, setIsLoading] = React.useState<boolean>(false); // Initially set to empty string

  // The initial list of pages
  const [initialPages, setInitialPages] = React.useState<any[]>([]); // Initially set to empty array

  // The column to sort by
  const [sortBy, setSortBy] = React.useState<string>(""); // Initially set to empty string
  const [scrollTop, setScrollTop] = React.useState<number>(0); // Initially set to empty string

  const [hasNextPage, setHasNextPage] = React.useState<boolean>(false);
  const [pageIndex, setPageIndex] = React.useState<number>(0);

  // The number of items to display per page
  const [pageSize, setPageSize] = React.useState<number>(40); // Initially set to 20

  // The total number of items
  const [totalItems, setTotalItems] = React.useState<number>(0); // Initially set to 0

  // The sorting order
  const [isDecending, setIsDecending] = React.useState<boolean>(false); // Initially set to false

  // Whether to show the filter panel
  const [showFilter, setShowFilter] = React.useState<boolean>(false); // Initially set to false

  // The column to filter by
  const [filterColumn, setFilterColumn] = React.useState<string>(""); // Initially set to empty string

  // The type of column to filter by
  const [filterColumnType, setFilterColumnType] = React.useState<string>(""); // Initially set to empty string

  // The filter details
  const [filterDetails, setFilterDetails] = React.useState<FilterDetail[]>([]); // Initially set to empty array

  // The taxonomy filter details
  const [taxonomyFilters, setTaxonomyFilters] = React.useState<FilterDetail[]>(
    []
  );

  // The filter details
  const [selectionDetails, setSelectionDetails] = React.useState<any | []>([]);
  // The filter details
  const [listId, setListId] = React.useState<string>("");
  const [currentUser, setCurrentUser] = React.useState<any>(null);
  const [viewId, setViewId] = React.useState<string>("");

  // Create an instance of the PagesService class
  const pagesService = new PagesService(context);

  // Get a unique id for the input field
  const inputId = useId("input");

  // Get the styles for the input field
  const inputStyles = useStyles();

  const subscribeIframeRef = React.useRef<HTMLIFrameElement>(null);

  React.useEffect(() => {
    const checkIframeUrl = () => {
      const iframe = subscribeIframeRef.current;
      if (iframe && iframe.contentWindow) {
        const currentUrl = iframe.contentWindow.location.href;

        if (
          currentUrl.indexOf("blank") === -1 &&
          currentUrl.indexOf(subscribeLink) === -1
        ) {
          setHideAlertMeDialog(true);
          setSelectionDetails([]);
          setPages([]);
          setTotalItems(0);
          setScrollTop(0);
          setHasNextPage(true);
          setPageIndex(0);

          getPages(catagory, columnInfos, pageSize, [], 0);
        }
      }
    };

    // Check the URL every 2 seconds
    const intervalId = setInterval(checkIframeUrl, 2000);

    // Clean up the interval on component unmount
    return () => clearInterval(intervalId);
  }, [subscribeLink, setHideAlertMeDialog, catagory, catagory]);

  /**
   * Resets the filters by clearing the checked items and
   * calling the applyFilters function with an empty filter detail.
   */
  const resetFilters = () => {
    // Clear the filter details
    setFilterDetails([]);
    setTaxonomyFilters([]);
    setTotalItems(0);
    setScrollTop(0);
    setHasNextPage(true);
    setPages([]);
    setPageIndex(0);

    // Clear the search text
    setSearchText("");

    // Call the fetchPages function with the default arguments
    fetchPages(
      pageSize,
      "Created",
      true,
      "",
      catagory,
      [],
      [],
      columnInfos,
      [],
      0
    );
  };

  /**
   * Fetches the paginated pages based on the given parameters.
   *
   * @param {number} [pageSizeAmount=pageSize] - The number of items per page. Defaults to the `pageSize` state variable.
   * @param {string} [sortBy="Created"] - The column to sort by. Defaults to "Created".
   * @param {boolean} [isSortedDescending=isDecending] - Whether to sort in descending order. Defaults to the `isDecending` state variable.
   * @param {string} [searchText=""] - The search text to filter by. Defaults to an empty string.
   * @param {string} [category=catagory] - The category to filter by. Defaults to the `catagory` state variable.
   * @param {FilterDetail[]} filterDetails - The filter details to apply.
   *
   * @return {Promise<void>} - A promise that resolves when the paginated pages are fetched.
   */
  const fetchPages = (
    pageSizeAmount = pageSize,
    sortBy = "Created",
    isSortedDescending = isDecending,
    searchText = "",
    category = catagory,
    filterDetails: FilterDetail[],
    taxonomyFilters: FilterDetail[] = [],
    columns: IColumnInfo[] = columnInfos,
    currentPages: any = pages,
    pageIndexPassed: number = pageIndex
  ) => {
    setIsLoading(true);
    setSelectionDetails([]);

    return (
      // Call the pagesService to fetch the filtered pages with non-taxonomy filters
      pagesService
        .getFilteredPages(
          sortBy,
          isSortedDescending,
          category as string,
          searchText,
          filterDetails, // Send only non-taxonomy filters to API
          columns,
          pageSizeAmount,
          pageIndexPassed
        )
        .then(
          (res: {
            pages: any;
            hasNextPage: boolean;
            nextPageIndex: number | null;
          }) => {
            let { pages, nextPageIndex } = res;
            // If there are taxonomy filters, apply them locally

            if (pages.length <= pageSize) setHasNextPage(false);
            else setHasNextPage(true);

            if (nextPageIndex) setPageIndex(nextPageIndex);
            if (taxonomyFilters.length > 0) {
              // Apply the taxonomy filters locally on the API results
              pages = pages.filter((item: any) => {
                // For each item, ensure that it matches every taxonomy filter
                return taxonomyFilters.every((taxonomyFilter) => {
                  // Check if the item has the taxonomy field and if it is an array
                  if (Array.isArray(item[taxonomyFilter.filterColumn])) {
                    // Ensure that the item matches at least one of the values in the filter
                    return taxonomyFilter.values.some((value) => {
                      // Check if the value matches any of the Label in the taxonomy items
                      return item[taxonomyFilter.filterColumn].some(
                        (taxonomyItem: { Label: string; TermGuid: string }) =>
                          taxonomyItem.Label === value
                      );
                    });
                  }
                  return false; // If the field does not exist or is not an array, filter it out
                });
              });
            }

            setHasNextPage(pages.length == pageSizeAmount);

            // Set the total number of items in the filtered pagesponse
            const finalPages = [...currentPages, ...pages];

            setPages(finalPages);
            setTotalItems(finalPages.length);

            setIsLoading(false);

            // Return the full filtered response
            return finalPages;
          }
        )
    );
  };

  /**
   * Fetches the pages from the given path and filter categories
   * and updates the state with the initial pages
   * @param path - The path to the SitePages library
   */
  const getPages = async (
    category: string | null,
    columns: IColumnInfo[],
    pageSize: number,
    currentPages: any,
    pageIndexPassed: number
  ): Promise<void> => {
    // Get the initial pages from the API

    const initialPagesFromApi = await fetchPages(
      pageSize,
      "Created",
      true,
      searchText,
      category,
      filterDetails,
      [],
      columns,
      currentPages,
      pageIndexPassed
    );

    // Update the state with the initial pages
    setInitialPages(initialPagesFromApi);
  };

  /**
   * Applies the given filter details to filter the pages
   *
   * @param {FilterDetail} filterDetail - The filter detail object containing the filter details
   */
  const applyFilters = (filterDetail: FilterDetail): void => {
    /**
     * Updates the current filter details state with the new filter detail,
     * or removes the filter detail if the values array is empty.
     *
     */
    let currentFilters: FilterDetail[] = filterDetails;
    let currentTaxonomyFilters: FilterDetail[] = taxonomyFilters;

    if (filterDetail.filterColumnType === "TaxonomyFieldTypeMulti") {
      if (filterDetail.values.length === 0) {
        currentTaxonomyFilters = taxonomyFilters.filter(
          (item) => item.filterColumn !== filterDetail.filterColumn
        );
      } else {
        currentTaxonomyFilters = [
          ...taxonomyFilters.filter(
            (item) => item.filterColumn !== filterDetail.filterColumn
          ),
          filterDetail,
        ];
      }
    } else {
      if (filterDetail.values.length === 0) {
        currentFilters = filterDetails.filter(
          (item) => item.filterColumn !== filterDetail.filterColumn
        );
      } else
        currentFilters = [
          ...filterDetails.filter(
            (item) => item.filterColumn !== filterDetail.filterColumn
          ),
          filterDetail,
        ];
    }
    setPageIndex(0);
    setFilterDetails(currentFilters);
    setTaxonomyFilters(currentTaxonomyFilters);

    fetchPages(
      pageSize, // Page size
      "Created", // Sorting criteria
      true, // Sorting order (ascending/descending)
      searchText, // Search text
      catagory, // Category (assuming this is another state or prop)
      currentFilters, // Updated filter details,
      currentTaxonomyFilters,
      columnInfos,
      [],
      0
    );
  };

  /**
   * Sort the pages list based on the specified column.
   *
   * @param {IColumn} column - The column to sort by.
   */
  const sortPages = (column: IColumn) => {
    // Set the sort by column state
    setSortBy(column.fieldName as string);

    // If the column is the same as the current sort by column, toggle the sort order
    if (column.fieldName === sortBy) {
      setIsDecending(!isDecending);
    } else {
      // Otherwise, set the sort order to descending
      setIsDecending(true);
    }

    // Fetch the pages list with the new sort criteria
    fetchPages(
      pageSize, // Page size
      column.fieldName, // Sorting criteria
      column.isSortedDescending, // Sorting order (ascending/descending)
      searchText, // Search text
      catagory, // Category (assuming this is another state or prop)
      filterDetails, // Filter details
      taxonomyFilters,
      columnInfos,
      [],
      0
    );
  };

  /**
   * Handles the search functionality by fetching pages with specified parameters.
   */
  const handleSearch = () => {
    fetchPages(
      pageSize, // Page size
      "Created", // Sorting criteria
      true, // Sorting order (ascending/descending)
      searchText, // Search text
      catagory, // Category
      filterDetails, // Filter details
      taxonomyFilters,
      columnInfos,
      [],
      0
    );
  };

  /**
   * Handles the change event of the page size dropdown.
   *
   * This function is triggered when the user selects a new page size from the dropdown.
   * It updates the page size state and calls the `handlePageChange` function to update
   * the paginated data.
   *
   * @function handlePageSizeChange
   * @memberof PagesList
   *
   * @param {any} e - The event object.
   * @return {void}
   */
  const handlePageSizeChange = (e: any) => {
    // Update the page size state
    setPageSize(e.target.value);
    setPages([]);
    setTotalItems(0);
    setScrollTop(0);

    setHasNextPage(true);
    setPageIndex(0);
    // Handle the page change with the new page size

    getPages(catagory, columnInfos, e.target.value, [], 0);
  };

  /**
   * Dismisses the filter panel.
   * Sets the showFilter state to false.
   *
   * @function dismissPanel
   * @memberof PagesList
   * @returns {void}
   */
  const dismissPanel = (): void => {
    setShowFilter(false);
  };

  const getColumns = async (selectedViewId: string) => {
    const columns = await pagesService.getColumns(selectedViewId);

    setColumnInfos(columns);

    return columns;
  };

  React.useEffect(() => {
    const handleEvent = (e: any) => {
      if (columnInfos.length > 0) {
        const selectedCategory = e.detail;

        if (selectedCategory && selectedCategory != "") {
          setCatagory(selectedCategory);

          getPages(selectedCategory, columnInfos, pageSize, [], 0);

          setSelectionDetails([]);
          setPageSize(pageSize);
        }
      }
    };

    pagesService.getListDetailsByName("Site Pages").then((res) => {
      setListId(res.Id);
    });

    window.addEventListener("catagorySelected", handleEvent);
  }, [columnInfos]);

  React.useEffect(() => {
    const fetchCurrentUser = async () => {
      try {
        const currentUserResponse = await context.spHttpClient.get(
          `${context.pageContext.web.absoluteUrl}/_api/web/currentuser`,
          SPHttpClient.configurations.v1
        );
        const userData = await currentUserResponse.json();
        setCurrentUser(userData);
      } catch (error) {
        console.error("Error fetching current user:", error);
      }
    };
    fetchCurrentUser();
  }, []);

  React.useEffect(() => {
    if (viewId !== selectedViewId) {
      setViewId(selectedViewId);
      getColumns(selectedViewId).then((col) => {
        if (catagory && catagory != "") {
          getPages(catagory, col, pageSize, [], 0);
        }
      });
    }
  }, [selectedViewId]);

  return (
    <div className="w-pageSize0 detail-display">
      {showFilter && (
        <FilterPanelComponent
          isOpen={showFilter}
          headerText="Filter Articles"
          applyFilters={applyFilters}
          dismissPanel={dismissPanel}
          selectedItems={
            [...filterDetails, ...taxonomyFilters].filter(
              (item) => item.filterColumn === filterColumn
            )[0] || { filterColumn: "", values: [] }
          }
          columnName={filterColumn}
          columnType={filterColumnType}
          pagesService={pagesService}
          data={initialPages}
          listId={listId}
        />
      )}
      <div className={`${styles.top}`}>
        <div
          className={`${styles["first-section"]} d-flex justify-content-between align-items-end py-2 px-2`}
        >
          <span className={`fs-4 ${styles["knowledgeText"]}`}>
            {catagory && <span className="">{catagory}</span>}
          </span>
          <div className={`${inputStyles.root} d-flex align-items-center me-2`}>
            <Input
              id={inputId}
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  handleSearch();
                }
              }}
              placeholder="Search"
            />
          </div>
        </div>

        <div
          className={`d-flex justify-content-between align-items-center fs-5 px-2 my-2`}
        >
          <span>Articles /</span>
          {totalItems > 0 ? (
            <div className="d-flex align-items-center">
              {selectionDetails && selectionDetails.length > 0 && (
                <DefaultButton
                  className="me-2"
                  onClick={() => {
                    toggleHideAlertMeDialog();
                  }}
                >
                  <span className="d-flex align-items-center">
                    <Icon iconName="Ringer" className="me-2" />
                    Alert Me
                  </span>
                </DefaultButton>
              )}
              {selectionDetails && selectionDetails.length > 0 && (
                <DefaultButton
                  className="me-2"
                  onClick={() => {
                    toggleHideManageAlertDialog();
                  }}
                >
                  <span className="d-flex align-items-center">
                    <Icon iconName="EditNote" className="me-2" />
                    Manage My Alerts
                  </span>
                </DefaultButton>
              )}
              {selectionDetails && selectionDetails.length > 0 && (
                <DefaultButton
                  className="me-2"
                  onClick={() => {
                    toggleHideFeedbackDialog();
                  }}
                >
                  <span className="d-flex align-items-center">
                    <Icon iconName="Feedback" className="me-2" />
                    Add Feedback
                  </span>
                </DefaultButton>
              )}
              {((filterDetails && filterDetails.length > 0) ||
                (taxonomyFilters && taxonomyFilters.length > 0)) && (
                <DefaultButton
                  onClick={() => {
                    resetFilters();
                  }}
                >
                  Clear
                </DefaultButton>
              )}
              <span className="ms-2 fs-6">Results ({totalItems})</span>
            </div>
          ) : (
            <span className="fs-6">No articles to display</span>
          )}
        </div>
      </div>

      {isLoading ? (
        <div style={{ textAlign: "center", minHeight: "300px" }}>
          <Spinner label="Articles are being loaded..." />
        </div>
      ) : (
        <div>
          <ReusableDetailList
            items={pages}
            context={context}
            columns={PagesColumns}
            columnInfos={columnInfos}
            currentUser={currentUser}
            setShowFilter={(column: IColumn, columnType: string) => {
              setShowFilter(!showFilter);
              setFilterColumn(column.fieldName as string);
              setFilterColumnType(columnType);
            }}
            updateSelection={(selection: Selection) => {
              setSelectionDetails(selection.getSelection());
            }}
            sortPages={sortPages}
            sortBy={sortBy}
            siteUrl={window.location.origin}
            isDecending={isDecending}
            loadMoreItems={() => {
              hasNextPage &&
                getPages(catagory, columnInfos, pageSize, pages, pageIndex);
            }}
            initialScrollTop={scrollTop}
            updateScrollPosition={(scrollTop: number) => {
              setScrollTop(scrollTop);
            }}
          />
        </div>
      )}
      <div className="d-flex justify-content-end">
        <div
          className="d-flex align-items-center my-1"
          style={{
            fontSize: "13px",
          }}
        >
          <div className="d-flex align-items-center me-3">
            <span className={`me-2 ${styles.blueText}`}>Items / Page </span>
            <select
              className="form-select"
              value={pageSize}
              onChange={handlePageSizeChange}
              name="pageSize"
              style={{
                width: 80,
                height: 35,
              }}
            >
              {pageSizeOption.map((pageSize) => {
                return (
                  <option key={pageSize} value={pageSize}>
                    {pageSize}
                  </option>
                );
              })}
            </select>
          </div>
        </div>
      </div>

      <Dialog
        hidden={hideFeedBackDialog}
        onDismiss={toggleHideFeedbackDialog}
        modalProps={{
          isBlocking: false,
        }}
        maxWidth="90vw"
        minWidth="60vw"
      >
        <ListForm
          articleId={
            selectionDetails[0] && selectionDetails[0].Article_x0020_ID
          }
          title={selectionDetails[0] && selectionDetails[0].Title}
          name={selectionDetails[0] && selectionDetails[0].FileLeafRef}
          link={
            selectionDetails[0] &&
            `${window.location.origin}${selectionDetails[0].FileRef}`
          }
          hideDialog={() => setHideFeedBackDialog(true)}
          pageService={pagesService}
          currentUser={currentUser}
          catagory={catagory}
          createdBy={selectionDetails[0] && selectionDetails[0].Author.Title}
          modifiedBy={selectionDetails[0] && selectionDetails[0].Editor.Title}
        />
      </Dialog>
      <Dialog
        hidden={hideAlertMeDialog}
        onDismiss={toggleHideAlertMeDialog}
        modalProps={{
          isBlocking: false,
        }}
        maxWidth="90vw"
        minWidth="60vw"
      >
        <iframe
          ref={subscribeIframeRef}
          src={`${
            context.pageContext.web.absoluteUrl
          }${subscribeLink}?List=${listId}&Id=${
            selectionDetails[0] && selectionDetails[0].Id
          }`}
          width="100%"
          height="600px"
          style={{ border: "none" }}
        ></iframe>

        <DialogFooter>
          <DefaultButton
            onClick={() => setHideAlertMeDialog(true)}
            text="Close"
          />
        </DialogFooter>
      </Dialog>
      <Dialog
        hidden={hideManageAlertDialog}
        onDismiss={toggleHideManageAlertDialog}
        modalProps={{
          isBlocking: false,
        }}
        maxWidth="90vw"
        minWidth="60vw"
      >
        <iframe
          src={`${context.pageContext.web.absoluteUrl}${alertLink}`}
          width="100%"
          height="600px"
          style={{ border: "none" }}
          id="alertFrame"
        ></iframe>

        <DialogFooter>
          <DefaultButton
            onClick={() => {
              setHideManageAlertDialog(true);
              getPages(catagory, columnInfos, pageSize, [], pageIndex);
            }}
            text="Close"
          />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default PagesList;
