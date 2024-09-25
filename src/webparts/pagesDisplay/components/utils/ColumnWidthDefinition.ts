export const WidthDefinition: {
  Title: string;
  Values: {
    MaxWidth: number;
    MinWidth: number;
  };
}[] = [
  {
    Title: "Article_x0020_ID",
    Values: {
      MinWidth: 60,
      MaxWidth: 80,
    },
  },
  {
    Title:
      "Title" ||
      "Name" ||
      "FileLeafRef" ||
      "LinkFilename" ||
      "LinkFilenameNoMenu",
    Values: {
      MinWidth: 400,
      MaxWidth: 1200,
    },
  },

  {
    Title: "Categories0",
    Values: {
      MinWidth: 200,
      MaxWidth: 800,
    },
  },

  {
    Title: "Modified",
    Values: {
      MinWidth: 200,
      MaxWidth: 200,
    },
  },
];
