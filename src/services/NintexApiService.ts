import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { INintexTask } from '../models/INintexTask';

export interface INintexUser {
  id: string;
  email: string;
  firstName?: string;
  lastName?: string;
}

export interface INintexAutoDelegation {
  id: string;
  userId: string;
  effectiveFrom: string;
  effectiveTo: string;
  message?: string;
  standIns: Array<{
    id: string;
    firstName?: string;
    lastName?: string;
    emails?: string[];
  }>;
  createdDate?: string;
  updatedDate?: string;
  fromUserDisplay?: string;
}

export class NintexApiService {
  private httpClient: HttpClient;
  private baseUrl: string;

  constructor(httpClient: HttpClient, baseUrl: string) {
    this.httpClient = httpClient;
    this.baseUrl = baseUrl ? baseUrl.replace(/\/$/, "") : "";
  }

  /**
   * Centralised error handler for Nintex API responses.
   * Shows a browser alert for 401 Unauthorized before throwing.
   */
  private async handleResponseError(response: HttpClientResponse, context: string): Promise<never> {
    const errorMsg = await response.text();
    if (response.status === 401) {
      alert(`Unauthorized: Your session has expired or you do not have permission to perform this action. Please contact your administrator.\n\nContext: ${context}`);
    }
    throw new Error(`${context}: ${errorMsg}`);
  }

  public async getPendingTasksForUser(userEmail: string, token: string): Promise<INintexTask[]> {
    if (!this.baseUrl) {
      throw new Error("Nintex API Base URL is not configured.");
    }

    const endpoint = `${this.baseUrl}/workflows/v2/tasks?assignee=${encodeURIComponent(userEmail)}&status=Active`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      }
    };

    const response: HttpClientResponse = await this.httpClient.get(endpoint, HttpClient.configurations.v1, options);
    if (!response.ok) {
      await this.handleResponseError(response, "Failed to fetch tasks from Nintex");
    }

    const data = await response.json();
    return data.tasks || [];
  }

  public async delegateTask(taskAssignmentId: string, delegateEmail: string, token: string): Promise<boolean> {
    const endpoint = `${this.baseUrl}/workflows/v2/tasks/${taskAssignmentId}/delegate`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        delegatees: [delegateEmail]
      })
    };

    const response: HttpClientResponse = await this.httpClient.post(endpoint, HttpClient.configurations.v1, options);
    
    if (!response.ok) {
      await this.handleResponseError(response, `Failed to delegate task ${taskAssignmentId}`);
    }

    return true;
  }



  public async getNintexUserId(email: string, token: string): Promise<string> {
    const endpoint = `${this.baseUrl}/tenants/v1/users/${encodeURIComponent(email)}`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      }
    };

    const response: HttpClientResponse = await this.httpClient.get(endpoint, HttpClient.configurations.v1, options);
    if (!response.ok) {
      await this.handleResponseError(response, `Failed to find Nintex user for email ${email}`);
    }

    const data = await response.json();
    if (!data.id) {
      throw new Error(`Nintex user ID not found in response for email ${email}.`);
    }

    return data.id;
  }

  public async createAutoDelegation(delegatorEmail: string, delegateEmail: string, startDateTime: Date, endDateTime: Date, token: string, message: string = ""): Promise<boolean> {
    const endpoint = `${this.baseUrl}/workflows/v2/tasks/autodelegations`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      // The tasks API (/workflows/v2) expects email addresses as user identifiers,
      // NOT the auth0| tenant user IDs returned by /tenants/v1/users.
      body: JSON.stringify({
        userId: delegatorEmail,
        effectiveFrom: startDateTime.toISOString(),
        effectiveTo: endDateTime.toISOString(),
        message: message,
        standIns: [
          {
            id: delegateEmail,
            type: "user"
          }
        ]
      })
    };

    const response: HttpClientResponse = await this.httpClient.post(endpoint, HttpClient.configurations.v1, options);
    
    if (!response.ok) {
      await this.handleResponseError(response, "Failed to create auto delegation");
    }

    return true;
  }

  public async deleteAutoDelegation(delegationId: string, token: string): Promise<boolean> {
    const endpoint = `${this.baseUrl}/workflows/v2/tasks/autodelegations/${delegationId}`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      },
      method: 'DELETE'
    };

    const response: HttpClientResponse = await this.httpClient.fetch(endpoint, HttpClient.configurations.v1, options);
    
    if (!response.ok) {
      await this.handleResponseError(response, "Failed to delete auto delegation");
    }

    return true;
  }

  public async updateAutoDelegation(delegationId: string, delegatorEmail: string, delegateEmail: string, startDateTime: Date, endDateTime: Date, token: string, message: string = ""): Promise<boolean> {
    const endpoint = `${this.baseUrl}/workflows/v2/tasks/autodelegations/${delegationId}`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      // The tasks API (/workflows/v2) expects email addresses as user identifiers,
      // NOT the auth0| tenant user IDs returned by /tenants/v1/users.
      body: JSON.stringify({
        userId: delegatorEmail,
        effectiveFrom: startDateTime.toISOString(),
        effectiveTo: endDateTime.toISOString(),
        message: message,
        standIns: [
          {
            id: delegateEmail,
            type: "user"
          }
        ]
      }),
      method: 'PUT'
    };

    const response: HttpClientResponse = await this.httpClient.fetch(endpoint, HttpClient.configurations.v1, options);
    
    if (!response.ok) {
      await this.handleResponseError(response, "Failed to update auto delegation");
    }

    return true;
  }

  public async listAutoDelegations(token: string, statusFilter: 'Active' | 'Scheduled' | 'Expired' | 'All' = 'All'): Promise<INintexAutoDelegation[]> {
    let endpoint = `${this.baseUrl}/workflows/v2/tasks/autodelegations`;
    
    const now = new Date();
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(now.getDate() - 30);

    const queryParams: string[] = [];

    if (statusFilter === 'Expired') {
      queryParams.push(`effectiveTo=${encodeURIComponent(thirtyDaysAgo.toISOString())}`);
    } else if (statusFilter === 'Active') {
      queryParams.push(`effectiveTo=${encodeURIComponent(now.toISOString())}`);
    }

    if (queryParams.length > 0) {
      endpoint += `?${queryParams.join('&')}`;
    }

    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      }
    };

    const response: HttpClientResponse = await this.httpClient.get(endpoint, HttpClient.configurations.v1, options);

    if (!response.ok) {
      await this.handleResponseError(response, "Failed to fetch auto delegations");
    }

    const data = await response.json();
    let delegations: INintexAutoDelegation[] = Array.isArray(data) ? data : data.taskAutoDelegations || data.data || [];

    const getItemStatus = (item: INintexAutoDelegation): 'Active' | 'Scheduled' | 'Expired' => {
      const f = new Date(item.effectiveFrom);
      const t = new Date(item.effectiveTo);
      if (now >= f && now <= t) return 'Active';
      if (now < f) return 'Scheduled';
      return 'Expired';
    };

    if (statusFilter !== 'All') {
      delegations = delegations.filter(item => {
        const status = getItemStatus(item);
        if (statusFilter === 'Expired') {
          if (status !== 'Expired') return false;
          const to = new Date(item.effectiveTo);
          return to >= thirtyDaysAgo;
        }
        return status === statusFilter;
      });
    } else {
      delegations = delegations.filter(item => {
        if (getItemStatus(item) === 'Expired') {
          const to = new Date(item.effectiveTo);
          return to >= thirtyDaysAgo;
        }
        return true;
      });
    }

    return delegations;
  }

  public async searchNintexUsers(filterText: string, token: string): Promise<INintexUser[]> {
    if (!filterText) return [];
    const endpoint = `${this.baseUrl}/tenants/v1/users?filter=${encodeURIComponent(filterText)}&limit=20`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      }
    };

    const response: HttpClientResponse = await this.httpClient.get(endpoint, HttpClient.configurations.v1, options);
    if (!response.ok) {
      if (response.status === 401) {
        alert("Unauthorized: Your session has expired or you do not have permission to search Nintex users. Please contact your administrator.");
      }
      console.error("Failed to fetch Nintex users", response.status);
      return [];
    }

    const data = await response.json();
    const rawUsers: INintexUser[] = Array.isArray(data) ? data : data.users || [];

    const lower = filterText.trim().toLowerCase();
    if (!lower) return rawUsers;
    const terms = lower.split(/\s+/).filter(t => t.length > 0);

    return rawUsers.filter((u: INintexUser) => {
      const fullName = `${u.firstName || ''} ${u.lastName || ''}`.trim().toLowerCase();
      const email = (u.email || '').toLowerCase();
      return terms.every(term => fullName.indexOf(term) !== -1 || email.indexOf(term) !== -1);
    });
  }

  public async getNintexUserById(id: string, token: string): Promise<INintexUser | undefined> {
    if (!id) return undefined;
    const endpoint = `${this.baseUrl}/tenants/v1/users/${encodeURIComponent(id)}`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      }
    };

    const response: HttpClientResponse = await this.httpClient.get(endpoint, HttpClient.configurations.v1, options);
    if (!response.ok) {
      if (response.status === 401) {
        alert("Unauthorized: Your session has expired or you do not have permission to fetch Nintex user details. Please contact your administrator.");
      }
      console.error(`Failed to fetch Nintex user by id ${id}`, response.status);
      return undefined;
    }

    return await response.json();
  }
}
