import { IColumn, Icon } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as React from "react";

const CellRender = (props: {
  columnName: string;
  columnType: string;
  item: any;
  context: WebPartContext;
}) => {
  const { columnName, item, columnType, context } = props;

  switch (columnType) {
    case "Text":
      return <div>{item[columnName]}</div>;
    case "DateTime":
      const date = new Date(item[columnName]);

      const optionsDate: any = {
        year: "numeric",
        month: "short",
        day: "numeric",
      };
      const formattedDate = date.toLocaleDateString("en-US", optionsDate);

      const optionsTime: any = {
        hour: "numeric",
        minute: "numeric",
        hour12: true,
      };
      const formattedTime = date.toLocaleTimeString("en-US", optionsTime);

      const formattedDateTime = `${formattedDate} ${formattedTime}`;
      return <div>{formattedDateTime}</div>;
    case "TaxonomyFieldTypeMulti":
      const taxMultiDetails =
        item[columnName] &&
        item[columnName].results
          .map((category: any) => category.Label)
          .join(", ");
      return <div>{taxMultiDetails}</div>;
    case "TaxonomyFieldType":
      const categories = item[columnName]
        .map((category: any) => category.Label)
        .join(", ");
      return <div>{categories}</div>;
    case "Number":
      return <div>{item[columnName]}</div>;
    case "User":
      return (
        item.FieldValuesAsText[columnName] && (
          <div>{item.FieldValuesAsText[columnName]}</div>
        )
      );
    case "URL":
      return (
        item[columnName] && (
          <a
            href={item[columnName].Url}
            target="_blank"
            title={`${window.origin}${item[columnName].Url}`}
          >
            {item[columnName].Description}
          </a>
        )
      );
    case "Computed":
      if (
        columnName === "Name" ||
        columnName === "FileLeafRef" ||
        columnName === "LinkFilename" ||
        columnName === "LinkFilenameNoMenu"
      ) {
        return (
          <a
            target="_blank"
            title={`${window.origin}${item.FieldValuesAsText.FileRef}`}
            style={{
              textDecoration: "underline",
              color: "blue",
              cursor: "pointer",
            }}
            onClick={(e) => {
              window.open(
                `${window.origin}${item.FieldValuesAsText.FileRef}`,
                "_blank"
              );
            }}
          >
            {item.FieldValuesAsText.FileLeafRef}
          </a>
        );
      } else {
        return (
          item[columnName] && (
            <a
              style={{
                textDecoration: "underline",
                color: "blue",
                cursor: "pointer",
              }}
              onClick={() => {
                window.open(
                  `${context.pageContext.web.absoluteUrl}/SitePages/${item[columnName]}`,
                  "_blank"
                );
              }}
              target="_blank"
              title={`${context.pageContext.web.absoluteUrl}/SitePages/${item[columnName]}`}
            >
              {item[columnName]}
            </a>
          )
        );
      }
    case "File":
      if (
        columnName === "FileLeafRef" ||
        columnName === "LinkFilename" ||
        columnName === "LinkFilenameNoMenu"
      ) {
        return (
          <a
            target="_blank"
            title={`${window.origin}${item.FieldValuesAsText.FileRef}`}
            style={{
              textDecoration: "underline",
              color: "blue",
              cursor: "pointer",
            }}
            onClick={(e) => {
              window.open(
                `${window.origin}${item.FieldValuesAsText.FileRef}`,
                "_blank"
              );
            }}
          >
            {item.FieldValuesAsText[columnName]}
          </a>
        );
      } else {
        return (
          item[columnName] && (
            <a
              target="_blank"
              title={`${context.pageContext.web.absoluteUrl}/SitePages/${item.FieldValuesAsText[columnName]}`}
              style={{
                textDecoration: "underline",
                color: "blue",
                cursor: "pointer",
              }}
              onClick={(e) => {
                window.open(
                  `${context.pageContext.web.absoluteUrl}/SitePages/${item.FieldValuesAsText[columnName]}`,
                  "_blank"
                );
              }}
            >
              {item[columnName]}
            </a>
          )
        );
      }
    default:
      return <div>{item[columnName]}</div>;
  }
};
const HeaderRender = (
  column: IColumn,
  columnType: string,
  onColumnClick: (column: IColumn) => void,
  setShowFilter: (column: IColumn, columnType: string) => void
): JSX.Element => {
  return (
    <div
      style={{
        display: "flex",
        alignItems: "start",
        justifyContent: "space-between",
        width: "100%", // Adjust padding as needed
        boxSizing: "border-box",
      }}
    >
      <span
        onClick={() => {
          if (column.fieldName !== "Categories0") {
            onColumnClick(column);
          }
        }}
        style={{
          cursor: "pointer",
        }}
      >
        {column.name}
      </span>

      <Icon
        iconName="Filter"
        onClick={() => setShowFilter(column, columnType)}
        style={{ cursor: "pointer" }}
      />
    </div>
  );
};

export { CellRender, HeaderRender };
