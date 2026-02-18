import * as React from "react";
import styles from "./AnnouncementDashboard.module.scss";
import type { IAnnouncementDashboardProps } from "./IAnnouncementDashboardProps";

import $ from "jquery";

import moment from "moment";
import { IAnnouncement } from "./IAnnouncement";
import { IAnnouncementDashboardState } from "./IAnnouncementDashboardState";

import { IColumn } from "@fluentui/react/lib/DetailsList";
import {
  TextField,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  CheckboxVisibility,
} from "@fluentui/react";
import { UI } from "../loc/il8n/ui";

const getAnnouncementColumns = (
  isArabic: boolean,
  t: any,
  onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => void,
): IColumn[] => {
  return [
    {
      key: "Title",
      name: t.HeaderTitle,
      fieldName: isArabic ? "Title_ar" : "Title_en",
      minWidth: 50,
      maxWidth: 150,
      isResizable: true,
      onColumnClick,
    },
    {
      key: "Description",
      name: t.HeaderDescription,
      fieldName: isArabic ? "Description_ar" : "Description_en",
      minWidth: 100,
      maxWidth: 300,
      isResizable: true,
      onColumnClick,

      isMultiline: true,
    },
    {
      key: "Category",
      name: t.HeaderCategory,
      fieldName: isArabic ? "category_ar" : "category_en",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
      onColumnClick,
    },
    {
      key: "Priority",
      name: t.HeaderPriority,
      fieldName: "Priority",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
      onColumnClick,
    },
    {
      key: "DueDate",
      name: t.HeaderDueDate,
      fieldName: "DueDate",
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      onColumnClick,
    },
    {
      key: "AssignedTo",
      name: t.HeaderAssignedTo,
      fieldName: "AssignedTo",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
      onColumnClick,
    },
  ];
};
const formatDate = (dateString: string): string => {
  return moment(dateString).format("M/DD/YYYY");
};
export interface IDetailListItem {
  Title_ar: string;
  Title_en: string;
  Description_ar: string;
  Description_en: string;
  Category: string;
  Priority: number;
  DueDate: string;
  AssignedTo: string;
}
export interface IDetailListState {
  items: IDetailListItem[];
}

export default class AnnouncementDashboard extends React.Component<
  IAnnouncementDashboardProps,
  IAnnouncementDashboardState
> {
  static siteUrl: string = "";
  public constructor(
    props: IAnnouncementDashboardProps,
    state: IAnnouncementDashboardState,
  ) {
    super(props);
    this.state = {
      allAnnouncements: [],
      announcements: [],
      columns: [],
      filterText: "",
      currentPage: 1,
      itemsPerPage: 5,
      sortField: "",
      sortDescending: false,
      isLoading: false,
      error: null,
    };
    AnnouncementDashboard.siteUrl = this.props.websiteUrl;

    this.updateState = this.updateState.bind(this);
    console.log("stage title from constructor");
  }

  private _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn,
  ): void => {
    const { columns, announcements } = this.state;
    const newColumns: IColumn[] = columns.slice();

    const currColumn: IColumn = newColumns.filter(
      (c) => c.key === column.key,
    )[0];

    // toggle sort
    const newIsDesc = currColumn.isSorted
      ? !currColumn.isSortedDescending
      : false;

    newColumns.forEach((col: IColumn) => {
      if (col.key === currColumn.key) {
        col.isSorted = true;
        col.isSortedDescending = newIsDesc;
      } else {
        col.isSorted = false;
        col.isSortedDescending = false;
      }
    });

    const newItems = copyAndSort(
      announcements,
      currColumn.fieldName!,
      newIsDesc,
    );

    this.setState({
      columns: newColumns,
      announcements: newItems,
      sortField: currColumn.fieldName!,
      sortDescending: newIsDesc,
    });
  };
  // Handle card layout sorting
  private _sortCardItems = (fieldName: string): void => {
    const { announcements, sortField, sortDescending } = this.state;

    // Toggle sort direction if clicking same field, otherwise reset to ascending
    const newIsDesc = sortField === fieldName ? !sortDescending : false;

    const newItems = copyAndSort(announcements, fieldName, newIsDesc);

    this.setState({
      announcements: newItems,
      sortField: fieldName,
      sortDescending: newIsDesc,
      currentPage: 1, // Reset to first page after sorting
    });
  };

  public componentDidMount() {
    this.fetchAnnouncements();
  }
  private _onChangeText = (
    _ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string,
  ): void => {
    if (!this.props.IsFiltering) return;
    const filterText = (newValue || "").toLowerCase();

    const isAr = this.props.Language === "AR";
    const titleField = isAr ? "Title_ar" : "Title_en";
    const descField = isAr ? "Description_ar" : "Description_en";
    const catField = isAr ? "category_ar" : "category_en";

    const source = this.state.allAnnouncements;

    const filtered = !filterText
      ? source
      : source.filter(
          (x: any) =>
            String(x[titleField] || "")
              .toLowerCase()
              .includes(filterText) ||
            String(x[descField] || "")
              .toLowerCase()
              .includes(filterText) ||
            String(x[catField] || "")
              .toLowerCase()
              .includes(filterText) ||
            String(x.AssignedTo || "")
              .toLowerCase()
              .includes(filterText) ||
            String(x.Priority ?? "")
              .toLowerCase()
              .includes(filterText) ||
            String(x.DueDate || "")
              .toLowerCase()
              .includes(filterText),
        );

    this.setState({
      filterText,
      announcements: filtered,
      currentPage: 1,
    });
  };
  private renderLayout(items: IAnnouncement[]): React.ReactNode {
    switch (this.props.Layout) {
      case "Card":
        return this.renderCardLayout(items);
      case "Compact":
        return this.renderDetailsListLayout(items, true);
      case "Table":
      default:
        return this.renderDetailsListLayout(items, false);
    }
  }
  private renderDetailsListLayout(
    items: IAnnouncement[],
    compact: boolean,
  ): React.ReactNode {
    const isAr = this.props.Language === "AR";

    return (
      <DetailsList
        key={`${isAr ? "rtl" : "ltr"}-${compact ? "compact" : "table"}`}
        className={[
          compact ? styles.detailsListCompact : styles.detailsList,
          isAr ? styles.rtl : styles.ltr,
        ].join(" ")}
        items={items}
        columns={
          compact
            ? this.state.columns.filter((c) => c.key !== "Description")
            : this.state.columns
        }
        setKey="set"
        layoutMode={DetailsListLayoutMode.fixedColumns}
        selectionMode={SelectionMode.single}
        checkboxVisibility={CheckboxVisibility.hidden}
        compact={compact}
      />
    );
  }
  private renderCardLayout(items: IAnnouncement[]): React.ReactNode {
    const isAr = this.props.Language === "AR";
    const t = isAr ? UI.AR : UI.EN;

    const titleField = isAr ? "Title_ar" : "Title_en";
    const descField = isAr ? "Description_ar" : "Description_en";
    const catField = isAr ? "category_ar" : "category_en";
    const { sortField, sortDescending } = this.state;

    return (
      <div>
        <div className={styles.cardSortControls}>
          <span className={styles.sortLabel}>{t.SortBy || "Sort by"}:</span>
          <button
            className={`${styles.sortButton} ${sortField === "Priority" ? styles.active : ""}`}
            onClick={() => this._sortCardItems("Priority")}
          >
            {t.HeaderPriority}
            {sortField === "Priority" && (
              <span className={styles.sortArrow}>
                {sortDescending ? " ▼" : " ▲"}
              </span>
            )}
          </button>
          <button
            className={`${styles.sortButton} ${sortField === "DueDate" ? styles.active : ""}`}
            onClick={() => this._sortCardItems("DueDate")}
          >
            {t.HeaderDueDate}
            {sortField === "DueDate" && (
              <span className={styles.sortArrow}>
                {sortDescending ? " ▼" : " ▲"}
              </span>
            )}
          </button>
          <button
            className={`${styles.sortButton} ${sortField === titleField ? styles.active : ""}`}
            onClick={() => this._sortCardItems(titleField)}
          >
            {t.HeaderTitle}
            {sortField === titleField && (
              <span className={styles.sortArrow}>
                {sortDescending ? " ▼" : " ▲"}
              </span>
            )}
          </button>
        </div>
        <section
          className={styles.cardsGrid}
          aria-label={t.HeaderTitle || "Announcements"}
        >
          {items.map((a, idx) => (
            <article key={idx} className={styles.card} tabIndex={0}>
              <header className={styles.cardHeader}>
                <h4 className={styles.cardTitle}>{(a as any)[titleField]}</h4>

                <div className={styles.badges}>
                  <span className={styles.badge}>{(a as any)[catField]}</span>
                  <span className={styles.badge}>
                    {t.HeaderPriority}: {a.Priority}
                  </span>
                </div>
              </header>

              <p className={styles.cardDescription}>{(a as any)[descField]}</p>

              <footer className={styles.cardFooter}>
                <span className={styles.metaItem}>
                  <strong>{t.HeaderDueDate}:</strong> {a.DueDate}
                </span>
                <span className={styles.metaItem}>
                  <strong>{t.HeaderAssignedTo}:</strong> {a.AssignedTo}
                </span>
              </footer>
            </article>
          ))}
        </section>
      </div>
    );
  }
  private renderPager(
    currentPage: number,
    totalPages: number,
    isAr: boolean,
    t: any,
  ): React.ReactNode {
    return (
      <nav
        className={styles.pager}
        aria-label={t.PaginationLabel || "Pagination"}
      >
        <PrimaryButton
          text={isAr ? "السابق" : "Previous"}
          disabled={currentPage === 1}
          onClick={() => this._getPage(currentPage - 1)}
        />

        <span className={styles.pagerText}>
          {t.Page} {currentPage} {t.of} {totalPages}
        </span>

        <PrimaryButton
          text={isAr ? "التالي" : "Next"}
          disabled={currentPage === totalPages}
          onClick={() => this._getPage(currentPage + 1)}
        />
      </nav>
    );
  }

  public componentDidUpdate(prevProps: IAnnouncementDashboardProps): void {
    // Re-fetch data if Items, Language, or ListName changes
    if (prevProps.Language !== this.props.Language) {
      const isAr = this.props.Language === "AR";
      const t = isAr ? UI.AR : UI.EN;

      this.setState({
        columns: getAnnouncementColumns(isAr, t, this._onColumnClick),
      });
    }
    if (
      prevProps.Items !== this.props.Items ||
      prevProps.Language !== this.props.Language ||
      prevProps.ListName !== this.props.ListName ||
      prevProps.Layout !== this.props.Layout
    ) {
      this.fetchAnnouncements();
    }
  }
  private fetchAnnouncements(): void {
    const apiUrl = `${AnnouncementDashboard.siteUrl}/sites/DevelopmentDemos/_api/web/lists/getbytitle('${this.props.ListName}')/items?$select=*,category/Title,category/name_en,category/name_ar,AssignedTo/Title,AssignedTo/EMail&$expand=category,AssignedTo`;

    $.ajax({
      url: apiUrl,
      type: "GET",
      headers: {
        Accept: "application/json;odata=verbose",
      },
      success: (resultData: any) => {
        console.log("AJAX success - data received:", resultData);
        let allAnnouncements: Array<IAnnouncement> = new Array<IAnnouncement>();
        resultData.d.results.map((item: any) => {
          allAnnouncements.push({
            Title_ar: item.Title_ar,
            Title_en: item.Title_en,
            Description_ar: item.Description_ar,
            Description_en: item.Description_en,
            category_en: item.category?.name_en || "Uncategorized",
            category_ar: item.category?.name_ar || "Uncategorized",
            Priority: item.Priority,
            DueDate: formatDate(item.DueDate),
            AssignedTo: item.AssignedTo?.Title || "Unassigned",
          });
        });
        console.log("Announcements processed:", allAnnouncements);
        const isAr = this.props.Language === "AR";
        const t = isAr ? UI.AR : UI.EN;
        this.setState({
          allAnnouncements,
          announcements: allAnnouncements,
          columns: getAnnouncementColumns(isAr, t, this._onColumnClick),
          currentPage: 1,
        });
      },
      error: (jqXHR: any, textStatus: string, errorThrown: string) => {
        console.error("AJAX ERROR");
        console.error("URL attempted:", apiUrl);
        console.error("Status:", textStatus);
        console.error("Error thrown:", errorThrown);
        console.error("HTTP Status Code:", jqXHR.status);
        console.error("Response text:", jqXHR.responseText);
      },
    });
  }

  public updateState() {
    this.setState({});
  }
  private _getPage = (pageNumber: number): void => {
    console.log("Page number clicked:", pageNumber);
    this.setState({ currentPage: pageNumber });
  };
  public render(): React.ReactElement<IAnnouncementDashboardProps> {
    const isAr = this.props.Language === "AR";
    const t = isAr ? UI.AR : UI.EN;

    const pageSize = this.props.Items;
    const totalItems = this.state.announcements.length;
    const totalPages = Math.max(1, Math.ceil(totalItems / pageSize));

    const currentPage = Math.min(this.state.currentPage, totalPages);
    const startIndex = (currentPage - 1) * pageSize;
    const pagedItems = this.state.announcements.slice(
      startIndex,
      startIndex + pageSize,
    );
    console.log("items per page", pageSize, "with total items", totalItems);
    return (
      <div
        key={this.props.Language}
        dir={this.props.Language === "AR" ? "rtl" : "ltr"}
        className={styles.announcementDashboard}
      >
        {this.props.IsFiltering && (
          <TextField
            label={t.FilterLabel}
            onChange={this._onChangeText}
            className={styles.controlStyles}
            value={this.state.filterText}
          />
        )}
        <div style={{ ["--adHeaderBg" as any]: this.props.color }}>
          {this.renderLayout(pagedItems)}
          {this.renderPager(currentPage, totalPages, isAr, t)}
        </div>
      </div>
    );
  }

  public componentWillUnmount() {
    console.log("component will unmount has been called");
  }
}
function copyAndSort<T>(
  items: T[],
  columnKey: string,
  isSortedDescending?: boolean,
): T[] {
  const key = columnKey as keyof T;

  return items.slice(0).sort((a: any, b: any) => {
    const av = a[key];
    const bv = b[key];

    // numeric if possible
    const an = Number(av);
    const bn = Number(bv);
    const bothNumbers = !Number.isNaN(an) && !Number.isNaN(bn);

    let cmp = 0;
    if (bothNumbers) cmp = an - bn;
    else
      cmp = String(av ?? "").localeCompare(String(bv ?? ""), undefined, {
        numeric: true,
      });

    return isSortedDescending ? -cmp : cmp;
  });
}
