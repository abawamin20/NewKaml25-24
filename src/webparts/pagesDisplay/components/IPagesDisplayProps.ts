import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPagesDisplayProps {
  context: WebPartContext;
  selectedViewId: string;
  feedbackPageUrl: string;
}
