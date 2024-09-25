import * as React from "react";
import { IColumn, IDetailsColumnProps } from "@fluentui/react";
import { Icon } from "@fluentui/react/lib/Icon";
import { IColumnInfo } from "./PagesService";
import { HeaderRender } from "../common/ColumnDetails";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";

// Global cache for subscription status
const subscriptionCache = new Map<string, boolean>();

/**
 * Returns an array of IColumn objects representing the columns of the PagesDetailsList component.
 *
 * @param {(column: IColumn) => void} onColumnClick - The function to call when a column is clicked.
 * @param {string} sortBy - The column to sort by.
 * @param {boolean} isDescending - Whether the sort order is descending.
 * @param {(column: IColumn) => void} setShowFilter - The function to set the showFilter state.
 * @return {IColumn[]} An array of IColumn objects representing the columns of the PagesDetailsList component.
 */
export const PagesColumns = (
  columns: IColumnInfo[],
  context: WebPartContext,
  currentUser: any,
  onColumnClick: (column: IColumn) => void,
  sortBy: string,
  isDescending: boolean,
  setShowFilter: (column: IColumn, columnType: string) => void
): IColumn[] => {
  const baseColumns = columns.map((column: IColumnInfo) => {
    return {
      key: column.InternalName,
      name: column.DisplayName,
      fieldName: column.InternalName,
      minWidth: column.MinWidth,
      maxWidth: column.MaxWidth,
      isRowHeader: true,
      isResizable: true,
      isPadded: true,
      isSorted: sortBy === column.InternalName,
      isSortedDescending: isDescending,
      onRenderHeader: (item: IDetailsColumnProps) =>
        HeaderRender(
          item.column,
          column.ColumnType,
          onColumnClick,
          setShowFilter
        ),
      onRender: column.OnRender
        ? column.OnRender
        : (item: any) => <div>{item[column.InternalName]}</div>,
    };
  });

  const statusColumn: IColumn = {
    key: "status",
    name: "Alert Status",
    fieldName: "status",
    minWidth: 80,
    maxWidth: 80,
    isRowHeader: false,
    isResizable: true,
    isPadded: true,
    onRender: (item: any) => (
      <SubscriptionStatus
        item={item}
        context={context}
        currentUser={currentUser}
      />
    ),
  };

  return [...baseColumns, statusColumn];
};
const SubscriptionStatus = ({
  item,
  context,
  currentUser,
}: {
  item: any;
  context: WebPartContext;
  currentUser: any;
}) => {
  const [subscribed, setSubscribed] = React.useState<boolean | null>(null);

  React.useEffect(() => {
    const itemKey = `SitePages_${item.Id}`;
    // Check if the status is already cached
    if (subscriptionCache.has(itemKey)) {
      setSubscribed(subscriptionCache.get(itemKey) as boolean);
      return; // Skip fetching if already cached
    }

    setSubscribed(false);
    const fetchSubscriptionStatus = async () => {
      try {
        const pageTitle = `Site Pages: ${item.FileLeafRef}`; // Ensure correct page title format
        const alertResponse = await context.spHttpClient.get(
          `${context.pageContext.web.absoluteUrl}/_api/web/alerts?$filter=UserId eq ${currentUser.Id} and Title eq '${pageTitle}'`,
          SPHttpClient.configurations.v1
        );
        const alertData = await alertResponse.json();

        // Set subscribed based on the presence of alerts
        const isSubscribed = alertData.value.length > 0;
        setSubscribed(isSubscribed);

        // Update the cache with the fetched status
        subscriptionCache.set(itemKey, isSubscribed);
      } catch (error) {
        console.error("Error fetching subscription status:", error);
        setSubscribed(false); // Default to false on error
      }
    };

    fetchSubscriptionStatus();
  }, [item.Title, context, currentUser.Id]); // Depend on specific props

  if (subscribed === null) {
    return <span>Loading...</span>; // Show loading indicator while fetching
  }

  return (
    <div className="status">
      {subscribed ? (
        <Icon iconName="RingerSolid" />
      ) : (
        <Icon iconName="Ringer" />
      )}
    </div>
  );
};
