import { WidthDefinition } from "./ColumnWidthDefinition";

const getColumnMinWidth = (columnInternalName: string): number => {
  const column = WidthDefinition.filter(
    (def) => def.Title === columnInternalName
  );
  return column.length > 0 ? column[0].Values.MinWidth : 100;
};
const getColumnMaxWidth = (columnInternalName: string): number => {
  const column = WidthDefinition.filter(
    (def) => def.Title === columnInternalName
  );
  return column.length > 0 ? column[0].Values.MaxWidth : 200;
};

export { getColumnMinWidth, getColumnMaxWidth };
