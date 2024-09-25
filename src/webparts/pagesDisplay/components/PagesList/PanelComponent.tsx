import * as React from "react";
import { Panel } from "@fluentui/react/lib/Panel";
import { Checkbox, Stack } from "@fluentui/react";
import { DefaultButton, PrimaryButton } from "@fluentui/react/lib/Button";
import { useId, Input } from "@fluentui/react-components";
import PagesService from "./PagesService";
import { FilterDetail } from "./PagesService";

export interface FilterOptions {
  key: string;
  text: string;
  value: string;
}

export interface ConstructedFilter {
  text: string;
  value: string;
}
const buttonStyles = { root: { marginRight: 8 } };

export const FilterPanelComponent = ({
  isOpen,
  dismissPanel,
  applyFilters,
  headerText,
  selectedItems,
  pagesService,
  columnName,
  columnType,
  data,
  listId,
}: {
  isOpen: boolean;
  dismissPanel: () => void;
  applyFilters: (filterDetail: FilterDetail) => void;
  headerText: string;
  selectedItems: FilterDetail;
  pagesService: PagesService;
  columnName: string;
  columnType: string;
  data: any[];
  listId: string;
}) => {
  const [checkedItems, setCheckedItems] =
    React.useState<FilterDetail>(selectedItems);
  const [searchText, setSearchText] = React.useState<string>("");
  const [filteredOptions, setFilteredOptions] = React.useState<FilterOptions[]>(
    []
  );
  const [options, setOptions] = React.useState<FilterOptions[]>([]);

  /**
   * Apply the filters by calling the applyFilters function
   * with the current checked items.
   */
  const apply = () => {
    // Create a filter detail object with the column name and values
    const filterDetail: FilterDetail = {
      filterColumn: columnName,
      filterColumnType: columnType,
      values: checkedItems.values,
    };

    // Call the applyFilters function with the filter detail
    applyFilters(filterDetail);
  };

  /**
   * Reset the filters by clearing the checked items and
   * calling the applyFilters function with an empty filter detail.
   */
  const resetFilters = () => {
    // Create a filter detail object with an empty array of values
    const filterDetail: FilterDetail = {
      filterColumn: columnName,
      filterColumnType: columnType,
      values: [],
    };

    // Update the checked items state with the filter detail
    setCheckedItems(filterDetail);

    // Call the applyFilters function with the filter detail
    applyFilters(filterDetail);
  };

  /**
   * Handles the search input change event by filtering the options
   * based on the search text and updating the filtered options state.
   */
  const handleSearch = () => {
    // Convert the search text to lowercase for case-insensitive search
    const lowercasedFilter = searchText.toLowerCase();

    // Filter the options based on the search text
    const filteredData = options.filter(
      (item) => item.text.toLowerCase().indexOf(lowercasedFilter) !== -1
    );

    // Update the filtered options state with the filtered data
    setFilteredOptions(filteredData);
  };

  /**
   * Constructs filter options for the given categories.
   *
   * @param categories - The list of categories to create filter options for.
   * @returns An array of FilterOptions for the categories.
   */
  const constructCategoryFilters = (
    categories: (number | string | ConstructedFilter)[]
  ) => {
    // Map categories to FilterOptions
    const updatedFilterCategories: FilterOptions[] = categories.map(
      (category: string | number | ConstructedFilter) => {
        if (typeof category === "string") {
          return {
            key: category,
            text: category,
            value: category,
          };
        } else if (typeof category === "number") {
          return {
            key: category.toString(), // Convert number to string
            text: category.toString(),
            value: category.toString(), // Ensure value is a string
          };
        } else {
          return {
            key: category.value,
            text: category.text,
            value: category.value,
          };
        }
      }
    );
    // Set the options state with the updated filter categories
    setOptions(updatedFilterCategories);

    // Set the filtered options state with all categories initially
    setFilteredOptions(updatedFilterCategories);

    return updatedFilterCategories;
  };

  /**
   * Effect hook that runs when the columnName prop changes.
   * This effect fetches distinct values for the columnName from the pagesService
   * and constructs category filters using the result.
   */
  React.useEffect(() => {
    // Fetch distinct values for the columnName from the pagesService
    pagesService
      .getDistinctValues(columnName, columnType, data)
      .then((res: (string | number | ConstructedFilter)[]) => {
        // Construct category filters using the result and update the options state

        constructCategoryFilters(res);
      });
  }, [columnName]);

  const onRenderFooterContent = () => (
    <div>
      <PrimaryButton onClick={apply} styles={buttonStyles}>
        Apply
      </PrimaryButton>
      <DefaultButton
        onClick={() => {
          resetFilters();
          setCheckedItems({
            filterColumn: columnName,
            filterColumnType: columnType,
            values: [],
          }); // Reset checked items to empty array
          setFilteredOptions(options); // Reset filtered options (if needed)
        }}
      >
        Clear
      </DefaultButton>
    </div>
  );

  const inputId = useId("input");

  return (
    <div>
      <Panel
        headerText={headerText}
        isOpen={isOpen}
        onDismiss={dismissPanel}
        onRenderFooterContent={onRenderFooterContent}
        closeButtonAriaLabel="Close"
        isFooterAtBottom={true}
      >
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
          style={{
            width: "100%",
            border: "1px solid black",
            marginBottom: "20px",
            paddingLeft: "10px",
          }}
        />

        <Stack tokens={{ childrenGap: 10 }}>
          {filteredOptions.map((option) => (
            <Checkbox
              key={option.key}
              label={
                isISODateString(option.text)
                  ? new Date(option.text).toLocaleDateString()
                  : option.text
              }
              checked={checkedItems.values.indexOf(option.value) !== -1}
              onChange={(ev, checked) => {
                if (checked) {
                  setCheckedItems({
                    filterColumn: columnName,
                    filterColumnType: columnType,
                    values: [...checkedItems.values, option.value],
                  });
                } else {
                  setCheckedItems({
                    filterColumn: columnName,
                    filterColumnType: columnType,
                    values: checkedItems.values.filter(
                      (item) => item !== option.value
                    ),
                  });
                }
              }}
            />
          ))}
        </Stack>
      </Panel>
    </div>
  );
};

// Function to check if a string is in ISO date format
function isISODateString(value: string): boolean {
  return /\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z?/.test(value);
}
