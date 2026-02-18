import { IAnnouncement } from "./IAnnouncement";

import { IColumn } from "@fluentui/react/lib/DetailsList";

export interface IAnnouncementDashboardState {
  announcements: Array<IAnnouncement>;
  //raw items
  allAnnouncements: Array<IAnnouncement>;
  columns: IColumn[];
  filterText: string;

  currentPage: number;
  itemsPerPage: number;
  sortField: string;
  sortDescending: boolean;

  isLoading: boolean;
  error: string | null;
}
