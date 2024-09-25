import * as React from "react";
import { useState } from "react";
import {
  TextField,
  PrimaryButton,
  DefaultButton,
  Stack,
  Label,
} from "@fluentui/react";
import PagesService from "../PagesList/PagesService";
import "./styles.css";

export interface IListFormProps {
  pageService: PagesService;
  articleId: string; // Passed from parent component
  title: string; // Passed from parent component
  name: string; // Passed from parent component
  link: string; // Passed from parent component
  catagory: string | null; // Passed from parent component
  hideDialog: () => void;
  currentUser: {
    Email: string;
    Title: string;
    Id: number;
  };
  createdBy: string;
  modifiedBy: string;
}

const CustomTextField = (props: any) => (
  <Stack
    horizontal
    verticalAlign="center"
    tokens={{ childrenGap: 10 }}
    style={{ width: props.width || "100%", marginBottom: "10px" }}
  >
    <Label styles={{ root: { maxWidth: "150px" } }}>{props.inputLabel}</Label>{" "}
    {/* Adjust label width */}
    <TextField
      value={props.value}
      type={props.type}
      styles={props.removeGrow ? undefined : { root: { flexGrow: 1 } }} // Make TextField expand
      readOnly
    />
  </Stack>
);

const ListForm: React.FunctionComponent<IListFormProps> = (props) => {
  const [feedbackComments, setFeedbackComments] = useState<string>("");

  const handleSubmit = async () => {
    const formData = {
      Article_x0020_ID: props.articleId,
      Title: props.title,
      Name: {
        Url: props.link,
        Description: props.name,
      },
      Link: {
        Url: props.link,
        Description: props.name,
      },
      FeedBackComments: feedbackComments,
      FeedBackProviderName: props.currentUser.Title,
      FeedBackProviderEmail: props.currentUser.Email,
      CreatedBy: props.createdBy,
      ModifiedBy: props.modifiedBy,
    };

    try {
      await props.pageService.createListItem(formData, "Feedbacks");
      alert("Feedback created successfully!");
      handleCancel(); // Clear the form and close the dialog
    } catch (error) {
      console.error("Error creating list item: ", error);
      alert("Failed to create item.");
    }
  };

  const handleCancel = () => {
    setFeedbackComments(""); // Clear the feedback field
    props.hideDialog(); // Close the dialog
  };

  return (
    <div>
      <h5
        style={{
          marginBottom: "30px",
        }}
      >
        Article Feedback
      </h5>

      <div className="intro">
        JCI Technical Knowledge Base(TKB) Article Feedback Submission
      </div>

      <div className="provider-detail">
        <span> Feedback Provider: </span> {props.currentUser.Title}
      </div>
      <div className="form-container">
        <Stack
          horizontal
          tokens={{ childrenGap: 15 }}
          style={{ width: "100%" }} // Add full width here
        >
          <CustomTextField
            inputLabel="Article Id"
            value={props.articleId}
            type="number"
            styles={{
              root: { flexGrow: 1 }, // Ensure the text field expands
              fieldGroup: { width: "100%" },
            }}
            width={"50%"}
            removeGrow={true}
          />

          <CustomTextField
            inputLabel="Knowledge Base"
            value={props.catagory ? props.catagory : ""}
            type="text"
            styles={{
              root: { flexGrow: 1 },
              fieldGroup: { width: "100%" },
            }}
          />
        </Stack>

        <CustomTextField inputLabel="Title" value={props.title} type="text" />

        <CustomTextField
          inputLabel="Link Name"
          value={props.name}
          type="text"
        />

        <CustomTextField
          inputLabel="Hypherlink"
          value={props.link}
          type="text"
        />
      </div>

      <TextField
        label="Feedback Comments"
        multiline
        rows={4}
        value={feedbackComments}
        onChange={(_, value) => setFeedbackComments(value || "")}
        style={{
          marginBottom: "10px",
        }}
      />

      <Stack
        horizontal
        tokens={{ childrenGap: 10 }}
        style={{ marginTop: "10px" }}
      >
        <PrimaryButton text="Submit Feedback" onClick={handleSubmit} />
        <DefaultButton text="Cancel" onClick={handleCancel} />
      </Stack>
    </div>
  );
};

export default ListForm;
