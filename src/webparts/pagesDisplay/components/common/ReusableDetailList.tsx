import * as React from "react";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  DetailsHeader,
  Selection,
  IDetailsListStyles,
} from "@fluentui/react/lib/DetailsList";
import { mergeStyles } from "@fluentui/react";
import "./styles.css";
import { IColumnInfo } from "../PagesList/PagesService";
import { WebPartContext } from "@microsoft/sp-webpart-base";

const gridStyles: Partial<IDetailsListStyles> = {
  root: {},
  headerWrapper: {},
};

const customHeaderClass = mergeStyles({
  backgroundColor: "#efefef",
  color: "white",
  paddingTop: 0,
  paddingBottom: 0,
  selectors: {
    "& .ms-DetailsHeader": {
      backgroundColor: "#0078d4",
      borderBottom: "1px solid #ccc",
    },
  },
});

export interface IReusableDetailListcomponentsProps {
  columns: (
    columns: IColumnInfo[],
    context: WebPartContext,
    currentUser: any,
    onColumnClick: any,
    sortBy: string,
    isDecending: boolean,
    setShowFilter: (column: IColumn, columnType: string) => void
  ) => IColumn[];
  columnInfos: IColumnInfo[];
  currentUser: any;
  context: WebPartContext;
  setShowFilter: (column: IColumn, columnType: string) => void;
  updateSelection: (selection: Selection) => void;
  items: any[];
  sortPages: (column: IColumn, isAscending: boolean) => void;
  sortBy: string;
  siteUrl: string;
  isDecending: boolean;
  loadMoreItems: () => void; // New prop to load more items
  initialScrollTop: number;
  updateScrollPosition: (scrollTop: number) => void;
}

export interface IReusableDetailListcomponentsState {}
export class ReusableDetailList extends React.Component<
  IReusableDetailListcomponentsProps,
  IReusableDetailListcomponentsState
> {
  private _selection: Selection;
  private containerRef: React.RefObject<HTMLDivElement>;

  constructor(props: IReusableDetailListcomponentsProps) {
    super(props);
    this._selection = new Selection({
      onSelectionChanged: () => {
        this.props.updateSelection(this._selection);
      },
      getKey: this._getKey,
    });

    this.state = {
      isLoading: false,
    };

    this.containerRef = React.createRef(); // Ref for the scrollable container
  }

  componentDidUpdate() {
    const { initialScrollTop } = this.props;

    if (this.containerRef.current) {
      // Restore scroll position on mount

      this.containerRef.current.scrollTop = initialScrollTop;
      this.containerRef.current.addEventListener("scroll", this.handleScroll);
    }
  }

  componentWillUnmount() {
    // Save scroll position before unmounting
    if (this.containerRef.current) {
      this.containerRef.current.removeEventListener(
        "scroll",
        this.handleScroll
      );
    }
  }

  handleScroll = () => {
    const container = this.containerRef.current;
    if (container) {
      const scrollTop = container.scrollTop;

      const scrollHeight = container.scrollHeight;
      const clientHeight = container.clientHeight;

      // Load more items when scrolled to the bottom
      if (scrollTop + clientHeight + 1 >= scrollHeight) {
        this.props.updateScrollPosition(scrollTop);
        this.props.loadMoreItems(); // Trigger loading more items from the parent
      }
    }
  };
  public render() {
    const {
      columnInfos,
      currentUser,
      context,
      columns,
      items,
      sortPages,
      sortBy,
      isDecending,
      setShowFilter,
    } = this.props;

    return (
      <div
        ref={this.containerRef} // Ref applied to the scrollable container
        style={{ maxHeight: "600px", overflowY: "auto" }}
      >
        <DetailsList
          styles={gridStyles}
          items={items}
          compact={true}
          columns={columns(
            columnInfos,
            context,
            currentUser,
            sortPages,
            sortBy,
            isDecending,
            setShowFilter
          )}
          selectionMode={SelectionMode.single}
          selection={this._selection}
          getKey={this._getKey}
          setKey="none"
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          onRenderDetailsHeader={this._onRenderDetailsHeader}
          onItemInvoked={this._onItemInvoked}
          className="detailList"
        />
      </div>
    );
  }

  private _getKey(item: any, index?: number): string {
    return item.key || index?.toString() || "";
  }

  private _onItemInvoked = (item: any): void => {
    window.open(`${this.props.siteUrl}${item.FileRef}`, "_blank");
  };

  private _onRenderDetailsHeader = (props: any) => {
    if (!props) {
      return null;
    }

    return (
      <DetailsHeader
        {...props}
        className="stickyHeader"
        styles={{ root: customHeaderClass }}
      />
    );
  };
}
