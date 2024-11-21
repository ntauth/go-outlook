package outlook

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strings"
	"time"

	"golang.org/x/oauth2"
)

const (
	// ClientVersion the current version of this sdk
	ClientVersion = "0.1.0"
	// DefaultBaseURL the root host url for the microsoft outlook api
	DefaultBaseURL = "https://graph.microsoft.com/v1.0"
	// DefaultOAuthTokenURL the url used to exchange a user's refreshToken for a usable accessToken
	DefaultOAuthTokenURL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
	// DefaultAuthScopes the set of permissions the client will request from the user
	DefaultAuthScopes = "mail.read calendars.read user.read offline_access"
	// DefaultQueryDateTimeFormat time format for the datetime query parameters used in outlook
	DefaultQueryDateTimeFormat = "2006-01-02T15:04:05Z"

	mediaType = "application/json"
)

var (
	// ErrNoDeltaLink error when our email paging fails to return a delta token at the end.
	ErrNoDeltaLink = errors.New("no delta link on response")

	// DefaultClient the http client that the sdk will use to make calls.
	DefaultClient = &http.Client{Timeout: time.Second * 60}

	// DefaultUserAgent the user agent to get passed in request headers on each call
	DefaultUserAgent = fmt.Sprintf("go-outlook/%s", ClientVersion)
)

// Client manages communication with microsoft's graph api, specifically for Mail and Calendar.
type Client struct {
	client      *http.Client
	baseURL     *url.URL
	userAgent   string
	mediaType   string
	tokenSource oauth2.TokenSource
}

// ClientOpt functions to configure options on a Client.
type ClientOpt func(*Client)

// SetClientMediaType returns a ClientOpt function which sets the clients mediaType.
func SetClientMediaType(mType string) ClientOpt {
	return func(c *Client) {
		c.mediaType = mType
	}
}

// SetClientTokenSource returns a ClientOpt function which sets the clients tokenSource.
func SetClientTokenSource(tokenSource oauth2.TokenSource) ClientOpt {
	return func(c *Client) {
		c.tokenSource = tokenSource
	}
}

// NewClient returns a new instance of a Client with the given options set.
func NewClient(opts ...ClientOpt) (*Client, error) {
	baseURL, err := url.Parse(DefaultBaseURL)
	if err != nil {
		return nil, err
	}
	client := &Client{
		client:    DefaultClient,
		baseURL:   baseURL,
		userAgent: DefaultUserAgent,
		mediaType: mediaType,
	}
	for _, opt := range opts {
		opt(client)
	}
	return client, nil
}

// SetMediaType fluent configuration of the client's mediaType.
func (client *Client) SetMediaType(mType string) *Client {
	client.mediaType = mType
	return client
}

// NewRequest creates a new request with some reasonable defaults based on the client.
func (client *Client) NewRequest(ctx context.Context, method, path string, body interface{}) (*http.Request, error) {
	var fullURL string
	pathURL, err := url.Parse(path)
	if err != nil {
		return nil, err
	}
	if pathURL.Hostname() != "" {
		fullURL = path
	} else {
		fullURL = fmt.Sprintf("%s%s", client.baseURL.String(), path)
	}

	encodedBody := new(bytes.Buffer)
	if body != nil {
		switch client.mediaType {
		case "application/json":
			if err := json.NewEncoder(encodedBody).Encode(body); err != nil {
				return nil, err
			}
		case "application/x-www-form-urlencoded":
			if v, ok := body.(url.Values); ok {
				bodyReader := strings.NewReader(v.Encode())
				if _, err := io.Copy(encodedBody, bodyReader); err != nil {
					return nil, err
				}
			} else {
				return nil, fmt.Errorf("body must be of type url.Values when Content-Type is set to application/x-www-form-urlencoded")
			}
		}
	}

	req, err := http.NewRequest(method, fullURL, encodedBody)
	if err != nil {
		return nil, err
	}

	req.Header.Add("Content-Type", client.mediaType)
	req.Header.Add("Accept", mediaType)
	req.Header.Add("User-Agent", client.userAgent)

	return req, nil
}

// Do executes the given http request and will bind the response body with v. Returns the http response as well as any error.
func (client *Client) Do(ctx context.Context, req *http.Request, v interface{}) (*http.Response, error) {
	req = req.WithContext(ctx)
	response, err := client.client.Do(req)
	if err != nil {
		return nil, err
	}

	defer func() {
		if closeErr := response.Body.Close(); closeErr != nil {
			err = closeErr
		}
	}()

	err = checkResponse(response)
	if err != nil {
		return response, err
	}

	if v != nil {
		if w, ok := v.(io.Writer); ok {
			_, err = io.Copy(w, response.Body)
			if err != nil {
				return response, err
			}
		} else {
			err = json.NewDecoder(response.Body).Decode(v)
			if err != nil {
				return response, err
			}
		}
	}

	return response, err
}

// NewSession returns a new instance of a Session using this client.
func (client *Client) NewSession() (*Session, error) {
	session, err := NewSession(client)
	if err != nil {
		return nil, err
	}

	return session, nil
}
