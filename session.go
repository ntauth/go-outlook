package outlook

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"net/url"
)

// Session manages communication to microsoft's graph api as an authenticated user.
type Session struct {
	client       *Client
	basePath     string
	accessToken  string
	refreshToken string
}

// NewSession returns a new instance of a Session.
func NewSession(client *Client) (*Session, error) {
	token, err := client.tokenSource.Token()
	if err != nil {
		return nil, err
	}

	session := &Session{
		client:       client,
		basePath:     "/me",
		accessToken:  token.AccessToken,
		refreshToken: token.RefreshToken,
	}

	return session, nil
}

func (session *Session) query(ctx context.Context, method, urlPath string, params map[string]interface{}, data interface{}, result interface{}) (*http.Response, error) {
	var queryString string
	if params != nil {
		queryString = createQueryString(params)
	}

	parsedBasePath, err := url.Parse(session.basePath)
	if err != nil {
		return nil, err
	}

	path := parsedBasePath.JoinPath(urlPath)
	if queryString != "" {
		path.RawQuery = queryString
	}

	req, err := session.client.NewRequest(ctx, method, path.String(), data)
	if err != nil {
		return nil, err
	}

	if session.accessToken == "" {
		return nil, ErrNoAccessToken
	}

	req.Header.Set("Authorization", fmt.Sprintf("Bearer %s", session.accessToken))

	// May want to detect failures due to invalid or expired tokens, then retry after attempting to refresh the token
	return session.client.Do(ctx, req, result)
}

// Get performs a get request to microsofts api with the underlying client and the sessions accessToken for authorization.
func (session *Session) Get(ctx context.Context, url string, params map[string]interface{}, result interface{}) (*http.Response, error) {
	return session.query(ctx, http.MethodGet, url, params, nil, result)
}

// Post performs a post request to microsofts api with the underlying client and the sessions accessToken for authorization.
func (session *Session) Post(ctx context.Context, url string, data interface{}, result interface{}) (*http.Response, error) {
	return session.query(ctx, http.MethodPost, url, nil, data, result)
}

// Patch performs a patch request to microsofts api with the underlying client and the sessions accessToken for authorization.
func (session *Session) Patch(ctx context.Context, url string, data interface{}, result interface{}) (*http.Response, error) {
	return session.query(ctx, http.MethodPatch, url, nil, data, result)
}

// Delete performs a delete request to microsofts api with the underlying client and the sessions accessToken for authorization.
func (session *Session) Delete(ctx context.Context, url string, params map[string]interface{}, result interface{}) (*http.Response, error) {
	return session.query(ctx, http.MethodDelete, url, params, nil, result)
}

// Calendars returns an instance of a CalendarService using this session.
func (session *Session) Calendars() *CalendarService {
	return NewCalendarService(session)
}

// Events returns an instance of a EventService using this session.
func (session *Session) Events() *EventService {
	return NewEventService(session)
}

// Folders returns an instance of a FolderService using this session.
func (session *Session) Folders() *FolderService {
	return NewFolderService(session)
}

// Messages returns an instance of a MessageService using this session.
func (session *Session) Messages() *MessageService {
	return NewMessageService(session)
}

func (s *Session) Send(ctx context.Context, message *Message) error {
	endpoint := "/sendMail"

	body := map[string]interface{}{
		"message": message,
	}

	// This method does not return any body, so we need to check for errors in the response
	resp, err := s.query(ctx, http.MethodPost, endpoint, nil, body, nil)
	if err != nil {
		return err
	}

	type APIError struct {
		Message string `json:"message"`
	}

	// Check for errors in the response
	if resp.StatusCode != http.StatusAccepted {
		var apiError APIError
		if err := json.NewDecoder(resp.Body).Decode(&apiError); err != nil {
			return fmt.Errorf("failed to send email: status %d: Failed to parse error response: %w", resp.StatusCode, err)
		}
		return fmt.Errorf("failed to send email: status %d: %s", resp.StatusCode, apiError.Message)
	}

	return nil
}
