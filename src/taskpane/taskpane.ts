/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

const awsmobile = {
  aws_project_region: "us-west-2",
  aws_cognito_identity_pool_id: "us-west-2:8b29f52b-27f0-4bc1-97ca-041c594b0091",
  aws_cognito_region: "us-west-2",
  aws_user_pools_id: "us-west-2_Gnmq0lww5",
  aws_user_pools_web_client_id: "2tjjblup0ki30td1plcdmoc9ff",
  oauth: {},
  aws_cognito_username_attributes: ["EMAIL"],
  aws_cognito_social_providers: [],
  aws_cognito_signup_attributes: ["EMAIL"],
  aws_cognito_mfa_configuration: "OFF",
  aws_cognito_mfa_types: ["SMS"],
  aws_cognito_password_protection_settings: {
    passwordPolicyMinLength: 8,
    passwordPolicyCharacters: [],
  },
  aws_cognito_verification_mechanisms: ["EMAIL"],
  aws_appsync_graphqlEndpoint: "https://n6kzyc46ljff5cfbhx5he7uhj4.appsync-api.us-west-2.amazonaws.com/graphql",
  aws_appsync_region: "us-west-2",
  aws_appsync_authenticationType: "AMAZON_COGNITO_USER_POOLS",
  aws_user_files_s3_bucket: "alucio-beacon-content104704-joseram",
  aws_user_files_s3_bucket_region: "us-west-2",
  aws_dynamodb_all_tables_region: "us-west-2",
  aws_dynamodb_table_schemas: [
    {
      tableName: "ObjectAudit-joseram",
      region: "us-west-2",
    },
  ],
  Auth: {
    identityPoolId: "us-west-2:8b29f52b-27f0-4bc1-97ca-041c594b0091", //REQUIRED - Amazon Cognito Identity Pool ID
    region: "us-west-2", // REQUIRED - Amazon Cognito Region
    userPoolId: "us-west-2_Gnmq0lww5", //OPTIONAL - Amazon Cognito User Pool ID
    userPoolWebClientId: "2tjjblup0ki30td1plcdmoc9ff", //OPTIONAL - Amazon Cognito Web Client ID
  },
  Storage: {
    AWSS3: {
      bucket: "alucio-beacon-content104704-joseram", //REQUIRED -  Amazon S3 bucket name
      region: "us-west-2", //OPTIONAL -  Amazon service region
    },
  },
};

type CreateDocumentFromS3UploadInput = {
  srcFilename: string;
  fileS3Key: string;
};

const createDocumentFromS3Upload = /* GraphQL */ `
  mutation CreateDocumentFromS3Upload($inputDoc: CreateDocumentFromS3UploadInput!) {
    createDocumentFromS3Upload(inputDoc: $inputDoc) {
      id
      tenantId
      documentId
      versionNumber
      srcFilename
      conversionStatus
      status
      srcFile {
        bucket
        region
        key
        url
      }
      srcHash
      srcSize
      pages {
        pageId
        number
        srcId
        srcHash
        isRequired
        speakerNotes
        linkedSlides
      }
      pageGroups {
        id
        pageIds
        name
        locked
      }
      type
      releaseNotes
      changeType
      labelValues {
        key
        value
      }
      customValues {
        fieldId
        values
      }
      title
      shortDescription
      longDescription
      owner
      expiresAt
      hasCopyright
      purpose
      canHideSlides
      distributable
      downloadable
      isInternalGenerated
      semVer {
        minor
        major
      }
      notificationScope
      selectedThumbnail
      publishedAt
      uploadedAt
      uploadedBy
      convertedArchiveKey
      convertedArchiveSize
      convertedFolderKey
      associatedFiles {
        id
        isDistributable
        isDefault
        type
        attachmentId
        status
        createdAt
        createdBy
      }
      editPermissions
      converterVersion
      createdAt
      createdBy
      updatedAt
      updatedBy
      integration {
        externalVersionId
        version
        timestamp
        srcFileHash
        srcDocumentHash
      }
      integrationType
      _version
      _deleted
      _lastChangedAt
    }
  }
`;

import { Amplify, Auth, Storage, API } from "aws-amplify";
const appsync = require("aws-appsync");
import gpl from "graphql-tag";
require("cross-fetch/polyfill");
Amplify.configure(awsmobile);

const docs = [
  {
    thumbnail: "../../assets/doc-1.png",
    title: "Document 1",
    description: "This is a document",
    url: "https://gist.githubusercontent.com/chibchombiano26/721ae21344f7ca71c3ed184c87aa1146/raw/027085a8d76a6818e2e4a68de2653fed2250a2eb/olympta",
  },
  {
    thumbnail: "../../assets/doc-2.png",
    title: "Document 1",
    description: "This is a document",
    url: "https://gist.githubusercontent.com/chibchombiano26/8aa719d5e61bde40d249a3e49c691591/raw/4e4c54c1c5ec3fa405dea085f9ae680de06975a0/document%25202",
  },
  {
    thumbnail: "../../assets/doc-3.png",
    title: "Document 1",
    description: "This is a document",
    url: "https://gist.githubusercontent.com/chibchombiano26/8aa719d5e61bde40d249a3e49c691591/raw/4e4c54c1c5ec3fa405dea085f9ae680de06975a0/document%25202",
  },
];

let graphClient = undefined;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("send-to-beacon").addEventListener("click", async () => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
        // @ts-ignore
        if (result.status == "succeeded") {
          const myFile = result.value;
          myFile.getSliceAsync(0, async (slice) => {
            // @ts-ignore
            if (slice.status == "succeeded") {
              const fileContentArry = slice.value.data;

              const fileContent = new Uint8Array(fileContentArry);
              const file = new File([fileContent], "myFile.pptx", { type: "data:attachment/powerpoint" });

              const result = await Storage.put("myFile.pptx", fileContent, {
                level: "private",
                contentType: "data:attachment/powerpoint",
                progressCallback(progress) {
                  console.log(progress);
                },
              });
              console.log(result);
              await createDocumentVersionFromS3(graphClient, result.key, "myFile.pptx");
            }
          });

          myFile.closeAsync();
        } else {
          // app.showNotification("Error:", result.error.message);
        }
      });
    });

    await signIn("alucioqa+josed-eclark@gmail.com", "TestUser321!");
    generateThumbnails();
  }
});

function generateThumbnails() {
  docs.forEach((doc) => {
    const node = document.createElement("div");
    const img = document.createElement("img");
    img.src = doc.thumbnail;
    img.onclick = () => {
      uploadFile(doc.url);
    };
    node.appendChild(img);
    document.getElementById("thumbnails").appendChild(node);
  });
}

export async function run() {
  /**
   * Insert your PowerPoint code here
   */
  const options: Office.SetSelectedDataOptions = { coercionType: Office.CoercionType.Text };

  await Office.context.document.setSelectedDataAsync(" ", options);
  await Office.context.document.setSelectedDataAsync("Hello World!", options);
}

const uploadFile = async (url: string) => {
  const blob = await fetch(url);
  const text = await blob.text();
  PowerPoint.createPresentation(text);
};

async function signIn(username, password) {
  try {
    const user = await Auth.signIn(username, password);
    graphClient = getGraphQLClientWithJWT(user.signInUserSession.idToken.jwtToken);
    console.log(user);
  } catch (error) {
    console.log("error signing in", error);
  }
}

async function uploadFromS3(fileName: string, srcFilename: string) {
  const updatedTodo = await API.graphql({
    query: createDocumentFromS3Upload,
    variables: {
      inputDoc: {
        fileS3Key: fileName,
        srcFilename: srcFilename,
      },
    },
  });

  console.log(updatedTodo);
}

function getGraphQLClientWithJWT(token) {
  const graphqlClient = new appsync.AWSAppSyncClient({
    url: awsmobile.aws_appsync_graphqlEndpoint,
    region: awsmobile.aws_appsync_region,
    auth: {
      type: "AMAZON_COGNITO_USER_POOLS",
      jwtToken: token,
    },
    disableOffline: true,
  });

  return graphqlClient;
}

const createDocumentVersionFromS3 = async (gpqlClient: any, fileS3Key: string, srcFilename: string) => {
  try {
    const result = (await gpqlClient.mutate({
      mutation: gpl(createDocumentFromS3Upload),
      variables: {
        inputDoc: {
          fileS3Key: fileS3Key,
          srcFilename: srcFilename,
        },
      },
    })) as { data: any };

    return result;
  } catch (e) {
    console.error(e);
    throw e;
  }
};
